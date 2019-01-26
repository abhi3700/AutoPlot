# Import modules
import xlwings as xw
import datetime as dt
import win32api
import matplotlib.pyplot as plt
import pandas as pd
import matplotlib.dates as mdates
from matplotlib.figure import Figure
from matplotlib.dates import MO, TU, WE, TH, FR, SA, SU
from matplotlib.lines import Line2D
from mpldatacursor import datacursor
# import plotly.plotly as py
# import plotly.graph_objs as go
# import os
# from pathlib import Path


def main():
    wb = xw.Book.caller()
    # wb.sheets[0].range("A1").value = "Hello xlwings!"		# test code

    #****************************************************************************************************************************************************************
    # Define sheets
    sht_asfe1_cp = wb.sheets['ASFE1-CP']
    sht_asfe1_er = wb.sheets['ASFE1-ER']
    sht_asfe1_plot_cp = wb.sheets['CP Plot']
    sht_asfe1_plot_er = wb.sheets['ER Plot']


    #****************************************************************************************************************************************************************
    # Fetch Dataframe for CP
    df_asfe1_cp = sht_asfe1_cp.range('A10').options(
        pd.DataFrame, header=1, index=False, expand='table'
        ).value											                # fetch the data from sheet- 'ASFE1-CP'
    df_asfe1_cp['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_asfe1_cp = df_asfe1_cp[["Date (MM/DD/YYYY)", "delta CP", "USL", "UCL", "Remarks"]]        # The final dataframe with required columns
    # sht_asfe1_plot_cp.range('A25').options(index=False).value = df_asfe1_cp   	    # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw CP Plot
    fig_asfe1_cp, ax_asfe1_cp = plt.subplots(1,1, figsize=(20,6))
    monthyearFmt_cp = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    ax_asfe1_cp.xaxis.set_major_formatter(monthyearFmt_cp)
    _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    ax_asfe1_cp.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    ax_asfe1_cp.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    plt.plot(df_asfe1_cp["Date (MM/DD/YYYY)"], df_asfe1_cp["delta CP"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs CP
    plt.plot(df_asfe1_cp["Date (MM/DD/YYYY)"], df_asfe1_cp["USL"], linestyle='-', color='#0000CD')        # plot date vs USL
    plt.plot(df_asfe1_cp["Date (MM/DD/YYYY)"], df_asfe1_cp["UCL"], linestyle='-', color='#FF1493')        # plot date vs UCL
    plt.xlabel('Date', fontsize=18)      # xlabel
    plt.ylabel('delta CP (no.s)', fontsize=18)     # ylabel
    # Custom Legends
    custom_lines_cp = [
        Line2D([0], [0], color='#FF7F50', lw=4),
        Line2D([0], [0], color='#0000CD', lw=4),
        Line2D([0], [0], color='#FF1493', lw=4)        
        ]
    ax_asfe1_cp.legend(custom_lines_cp, ['CP', 'USL', 'UCL'], fontsize=11, loc='upper right')  
    lines_cp = ax_asfe1_cp.plot(df_asfe1_cp["Date (MM/DD/YYYY)"], df_asfe1_cp["delta CP"], visible=False)
    datacursor(lines_cp, hover=True, point_labels=df_asfe1_cp['Remarks'])
    # plt.show()
    # sht_asfe1_plot_cp.activate()
    # pic_cp = plt.show()
    # plt.show('ASFE1_CP_Plot', left=xw.Range('A1').left, top=xw.Range('A1').top)      # this would activate hover 
    sht_asfe1_plot_cp.pictures.add(fig_asfe1_cp, name= "ASFE1_CP_Plot", update= True)
    # sht_asfe1_plot_cp.pictures.add(pic_cp, name= "ASFE1_CP_Plot", update= True)


    #****************************************************************************************************************************************************************
    # Fetch Dataframe for ER   
    # data_folder = Path(os.getcwd())
    # file_to_open = data_folder / "ASH09_QC_LOG_BOOK.xlsm"
    # excel_file = pd.ExcelFile(file_to_open)

    excel_file = pd.ExcelFile("H:\\excel\\dryetch\\Excel-office\\macro_enabled_logbooks\\ASH09_QC_LOG_BOOK\\ASH09_QC_LOG_BOOK.xlsm")
    # excel_file = pd.ExcelFile("\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\ASH_09_10_LOG_BOOK\\ASH09_QC_LOG_BOOK_macro\\ASH09_QC_LOG_BOOK.xlsm")
    df_asfe1_er = excel_file.parse('ASFE1-ER', skiprows=8)                            # copy a sheet and paste into another sheet and skiprows 8
    df_asfe1_er['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_asfe1_er = df_asfe1_er[["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "LSL", "LCL", "UCL", "Remarks", "% Uni UCL"]]             # The final Dataframe with 5 columns for plot: x-1, y-4
    df_asfe1_er = df_asfe1_er.dropna()                                              # dropping rows where at least one element is missing
    # sht_asfe1_plot_er.range('A28').options(index=False).value = df_asfe1_er        # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw ER PLot
    fig_asfe1_er, ax_asfe1_er = plt.subplots(1,1, figsize=(20,6))
    monthyearFmt_er = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    ax_asfe1_er.xaxis.set_major_formatter(monthyearFmt_er)
    _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    ax_asfe1_er.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    ax_asfe1_er.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    plt.plot(df_asfe1_er["Date (MM/DD/YYYY)"], df_asfe1_er["Etch Rate (A/Min)"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    plt.plot(df_asfe1_er["Date (MM/DD/YYYY)"], df_asfe1_er["UCL"], linestyle='-', color='#FF1493')        # plot date vs UCL
    plt.plot(df_asfe1_er["Date (MM/DD/YYYY)"], df_asfe1_er["LSL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    plt.plot(df_asfe1_er["Date (MM/DD/YYYY)"], df_asfe1_er["LCL"], linestyle='-', color='#FF1493')        # plot date vs LCL 
    plt.xlabel('Date', fontsize=18)      # xlabel
    plt.ylabel('Etch Rate (A/min)', fontsize=18)     # ylabel
    # Custom Legends
    custom_lines_asfe1_er = [
        Line2D([0], [0], color='#FF7F50', lw=4),
        Line2D([0], [0], color='#FF1493', lw=4),
        Line2D([0], [0], color='#0000CD', lw=4),
        Line2D([0], [0], color='#FF1493', lw=4)        
        ]
    ax_asfe1_er.legend(custom_lines_asfe1_er, ['ER', 'UCL', 'LSL', 'LCL'], fontsize=11, loc='upper right') 
    lines_er = ax_asfe1_er.plot(df_asfe1_er["Date (MM/DD/YYYY)"], df_asfe1_er["Etch Rate (A/Min)"], visible=False)
    datacursor(lines_er, hover=True, point_labels=df_asfe1_er['Remarks'])
    # plt.show()        # shows 2 figures in different windows
    sht_asfe1_plot_er.pictures.add(fig_asfe1_er, name= "ASFE1_ER_Plot", update= True)

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw Unif PLot
    fig_asfe1_unif, ax_asfe1_unif = plt.subplots(1,1, figsize=(20,6))
    monthyearFmt_unif = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    ax_asfe1_unif.xaxis.set_major_formatter(monthyearFmt_unif)
    _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    ax_asfe1_unif.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    ax_asfe1_unif.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    plt.plot(df_asfe1_er["Date (MM/DD/YYYY)"], df_asfe1_er["% Uni"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    plt.plot(df_asfe1_er["Date (MM/DD/YYYY)"], df_asfe1_er["% Uni UCL"], linestyle='-', color='#FF1493')        # plot date vs UCL
    plt.xlabel('Date', fontsize=18)      # xlabel
    plt.ylabel('Uniformity (%)', fontsize=18)     # ylabel
    # Custom Legends
    custom_lines_asfe1_unif = [
        Line2D([0], [0], color='#FF7F50', lw=4),
        Line2D([0], [0], color='#FF1493', lw=4),
        ]
    ax_asfe1_unif.legend(custom_lines_asfe1_unif, ['Unif', 'UCL'], fontsize=11, loc='upper right') 
    lines_unif = ax_asfe1_unif.plot(df_asfe1_er["Date (MM/DD/YYYY)"], df_asfe1_er["% Uni"], visible=False)
    datacursor(lines_unif, hover=True, point_labels=df_asfe1_er['Remarks'])
    # plt.show()        # shows 2 figures in different windows
    sht_asfe1_plot_er.pictures.add(fig_asfe1_unif, name= "ASFE1_Unif_Plot", update= True)


#--------------------------------------------------------------------------------------------------------------------------------
# User Defined Functions (UDFs)
#--------------------------------------------------------------------------------------------------------------------------------
@xw.func
def hello(name):
    return "hello {0}".format(name)



