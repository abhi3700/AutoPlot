# Import modules
import xlwings as xw
import datetime as dt
import win32api
import matplotlib.pyplot as plt
import pandas as pd
import matplotlib.dates as mdates
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
    sht_repl1a_cp = wb.sheets['REPL1A-CP']
    sht_repl1a_er_nit = wb.sheets['REPL1A-ERNit']
    sht_repl1a_er_poly = wb.sheets['REPL1A-ERPoly']
    sht_repl1a_plot_cp = wb.sheets['CP Plot']
    sht_repl1a_plot_er_sin = wb.sheets['SiN Plot']
    sht_repl1a_plot_er_poly = wb.sheets['Poly Plot']

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for CP
    df_repl1a_cp = sht_repl1a_cp.range('A9').options(
        pd.DataFrame, header=1, index=False, expand='table'
        ).value											                # fetch the data from sheet- 'ASBE1-CP'
    df_repl1a_cp['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_repl1a_cp = df_repl1a_cp[["Date (MM/DD/YYYY)", "delta CP", "USL", "Remarks"]]        # The final dataframe with required columns
    # sht_repl1a_plot_cp.range('A25').options(index=False).value = df_repl1a_cp   	    # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw CP Plot
    fig_repl1a_cp, fig_repl1a_cp = plt.subplots(1,1, figsize=(20,6))
    monthyearFmt_cp = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    fig_repl1a_cp.xaxis.set_major_formatter(monthyearFmt_cp)
    _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # fig_repl1a_cp.xaxis.set_major_locator(mdates.MonthLocator())          # set ticks after every Month
    fig_repl1a_cp.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    fig_repl1a_cp.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    plt.plot(df_repl1a_cp["Date (MM/DD/YYYY)"], df_repl1a_cp["delta CP"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs CP
    plt.plot(df_repl1a_cp["Date (MM/DD/YYYY)"], df_repl1a_cp["USL"], linestyle='-', color='#0000CD')        # plot date vs USL
    plt.xlabel('Date', fontsize=18)      # xlabel
    plt.ylabel('delta CP (no.s)', fontsize=18)     # ylabel
    # Custom Legends
    custom_lines_cp = [
        Line2D([0], [0], color='#FF7F50', lw=4),
        Line2D([0], [0], color='#0000CD', lw=4),
        ]
    fig_repl1a_cp.legend(custom_lines_cp, ['CP', 'USL'], fontsize=11, loc='upper right')  
    lines_cp = fig_repl1a_cp.plot(df_repl1a_cp["Date (MM/DD/YYYY)"], df_repl1a_cp["delta CP"], visible=False)
    datacursor(lines_cp, hover=True, point_labels=df_repl1a_cp['Remarks'])
    # plt.show()
    # sht_repl1a_plot_cp.activate()
    # pic_cp = plt.show()
    # plt.show('ASBE1_CP_Plot', left=xw.Range('A1').left, top=xw.Range('A1').top)      # this would activate hover 
    # sht_repl1a_plot_cp.pictures.add(pic_cp, name= "ASFE1_CP_Plot", update= True)
    sht_repl1a_plot_cp.pictures.add(fig_repl1a_cp, name= "REPL1A_CP_Plot", update= True)


    #****************************************************************************************************************************************************************
    # Fetch Dataframe for NIT  
    # data_folder = Path(os.getcwd())
    # file_to_open = data_folder / "ASH09_QC_LOG_BOOK.xlsm"
    # excel_file = pd.ExcelFile(file_to_open)

    # excel_file_sht_nit = pd.ExcelFile("H:\\excel\\dryetch\\Excel-office\\macro_enabled_logbooks\\CNT01_QC_LOG_BOOK_Ch_A\\CNT01_QC_LOG_BOOK_Ch_A.xlsm")
    excel_file_sht_nit = pd.ExcelFile("\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\CNT_01_LOG_BOOK\\CNT01_QC_LOG_BOOK_Ch_A_macro\\CNT01_QC_LOG_BOOK_Ch_A.xlsm")
    df_repl1a_er_nit = excel_file_sht_nit.parse('REPL1A-ERNit', skiprows=9)                            # copy a sheet and paste into another sheet and skiprows 9
    
    df_repl1a_er_nit = df_repl1a_er_nit[["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uniformity", "LSL", "USL", "LCL", "UCL", "Remarks", "% Uni USL", "% Uni UCL"]]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_repl1a_er_nit['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_repl1a_er_nit = df_repl1a_er_nit.dropna()                                              # dropping rows where at least one element is missing
    # sht_repl1a_plot_er_sin.range('A28').options(index=False).value = df_repl1a_er_nit        # show the dataframe values into sheet- 'CP Plot'
    
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw NIT ER PLot
    fig_repl1a_er_sin, ax_repl1a_er_nit = plt.subplots(1,1, figsize=(20,6))
    monthyearFmt_er_nit = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    ax_repl1a_er_nit.xaxis.set_major_formatter(monthyearFmt_er_nit)
    _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # ax_repl1a_er_nit.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    ax_repl1a_er_nit.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    ax_repl1a_er_nit.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    plt.plot(df_repl1a_er_nit["Date (MM/DD/YYYY)"], df_repl1a_er_nit["Etch Rate (A/Min)"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    plt.plot(df_repl1a_er_nit["Date (MM/DD/YYYY)"], df_repl1a_er_nit["USL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    plt.plot(df_repl1a_er_nit["Date (MM/DD/YYYY)"], df_repl1a_er_nit["LSL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    plt.plot(df_repl1a_er_nit["Date (MM/DD/YYYY)"], df_repl1a_er_nit["UCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    plt.plot(df_repl1a_er_nit["Date (MM/DD/YYYY)"], df_repl1a_er_nit["LCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    plt.xlabel('Date', fontsize=18)      # xlabel
    plt.ylabel('SiN ER (A/min)', fontsize=18)     # ylabel
    # Custom Legends
    custom_lines_er_nit = [
        Line2D([0], [0], color='#FF7F50', lw=4),	# ER
        Line2D([0], [0], color='#0000CD', lw=4),	# USL
        Line2D([0], [0], color='#0000CD', lw=4),	# LSL
        Line2D([0], [0], color='#FF1493', lw=4),	# UCL
        Line2D([0], [0], color='#FF1493', lw=4)		# LCL        
        ]
    ax_repl1a_er_nit.legend(custom_lines_er_nit, ['ER', 'USL', 'LSL', 'UCL', 'LCL'], fontsize=11, loc='upper right') 
    lines_er_nit = ax_repl1a_er_nit.plot(df_repl1a_er_nit["Date (MM/DD/YYYY)"], df_repl1a_er_nit["Etch Rate (A/Min)"], visible=False)
    datacursor(lines_er_nit, hover=True, point_labels=df_repl1a_er_nit['Remarks'])
    # plt.show()        # shows 2 figures in different windows
    sht_repl1a_plot_er_sin.pictures.add(fig_repl1a_er_sin, name= "REPL1A_NIT_ER_Plot", update= True)

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
	# Draw NIT Unif PLot
    fig_repl1a_unif_nit, ax_repl1a_unif_nit = plt.subplots(1,1, figsize=(20,6))
    monthyearFmt_unif_nit = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    ax_repl1a_unif_nit.xaxis.set_major_formatter(monthyearFmt_unif_nit)
    _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # ax_repl1a_unif_nit.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    ax_repl1a_unif_nit.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    ax_repl1a_unif_nit.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    plt.plot(df_repl1a_er_nit["Date (MM/DD/YYYY)"], df_repl1a_er_nit["% Uniformity"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    plt.plot(df_repl1a_er_nit["Date (MM/DD/YYYY)"], df_repl1a_er_nit["% Uni USL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    plt.plot(df_repl1a_er_nit["Date (MM/DD/YYYY)"], df_repl1a_er_nit["% Uni UCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    plt.xlabel('Date', fontsize=18)      # xlabel
    plt.ylabel('SiN Unif (%)', fontsize=18)     # ylabel
    # Custom Legends
    custom_lines_unif_nit = [
        Line2D([0], [0], color='#FF7F50', lw=4),	# ER
        Line2D([0], [0], color='#0000CD', lw=4),	# USL
        Line2D([0], [0], color='#FF1493', lw=4),	# UCL
        ]
    ax_repl1a_unif_nit.legend(custom_lines_unif_nit, ['Unif', 'USL', 'UCL'], fontsize=11, loc='upper right') 
    lines_unif_nit = ax_repl1a_unif_nit.plot(df_repl1a_er_nit["Date (MM/DD/YYYY)"], df_repl1a_er_nit["% Uniformity"], visible=False)
    datacursor(lines_unif_nit, hover=True, point_labels=df_repl1a_er_nit['Remarks'])
    # plt.show()        # shows 2 figures in different windows
    sht_repl1a_plot_er_sin.pictures.add(fig_repl1a_unif_nit, name= "REPL1A_NIT_UNIF_Plot", update= True)


    #****************************************************************************************************************************************************************
    # Fetch Dataframe for POLY   
    # data_folder = Path(os.getcwd())
    # file_to_open = data_folder / "ASH09_QC_LOG_BOOK.xlsm"
    # excel_file = pd.ExcelFile(file_to_open)

    # excel_file_sht_poly = pd.ExcelFile("H:\\excel\\dryetch\\Excel-office\\macro_enabled_logbooks\\CNT01_QC_LOG_BOOK_Ch_A\\CNT01_QC_LOG_BOOK_Ch_A.xlsm")
    excel_file_sht_poly = pd.ExcelFile("\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\CNT_01_LOG_BOOK\\CNT01_QC_LOG_BOOK_Ch_A_macro\\CNT01_QC_LOG_BOOK_Ch_A.xlsm")
    df_repl1a_er_poly = excel_file_sht_poly.parse('REPL1A-ERPoly', skiprows=9)                            # copy a sheet and paste into another sheet and skiprows 9
    
    df_repl1a_er_poly = df_repl1a_er_poly[["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uniformity", "LSL", "USL", "LCL", "UCL", "Remarks", "% Uni USL", "% Uni UCL"]]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_repl1a_er_poly['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_repl1a_er_poly = df_repl1a_er_poly.dropna()                                              # dropping rows where at least one element is missing
    # sht_repl1a_plot_er_poly.range('A28').options(index=False).value = df_repl1a_er_poly        # show the dataframe values into sheet- 'CP Plot'
    
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw POLY ER PLot
    fig_repl1a_er_poly, ax_repl1a_er_poly = plt.subplots(1,1, figsize=(20,6))
    monthyearFmt_er_poly = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    ax_repl1a_er_poly.xaxis.set_major_formatter(monthyearFmt_er_poly)
    _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # ax_repl1a_er_poly.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    ax_repl1a_er_poly.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    ax_repl1a_er_poly.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    plt.plot(df_repl1a_er_poly["Date (MM/DD/YYYY)"], df_repl1a_er_poly["Etch Rate (A/Min)"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    plt.plot(df_repl1a_er_poly["Date (MM/DD/YYYY)"], df_repl1a_er_poly["USL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    plt.plot(df_repl1a_er_poly["Date (MM/DD/YYYY)"], df_repl1a_er_poly["LSL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    plt.plot(df_repl1a_er_poly["Date (MM/DD/YYYY)"], df_repl1a_er_poly["UCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    plt.plot(df_repl1a_er_poly["Date (MM/DD/YYYY)"], df_repl1a_er_poly["LCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    plt.xlabel('Date', fontsize=18)      # xlabel
    plt.ylabel('Poly ER (A/min)', fontsize=18)     # ylabel
    # Custom Legends
    custom_lines_er_poly = [
        Line2D([0], [0], color='#FF7F50', lw=4),	# ER
        Line2D([0], [0], color='#0000CD', lw=4),	# USL
        Line2D([0], [0], color='#0000CD', lw=4),	# LSL
        Line2D([0], [0], color='#FF1493', lw=4),	# UCL
        Line2D([0], [0], color='#FF1493', lw=4)		# LCL        
        ]
    ax_repl1a_er_poly.legend(custom_lines_er_poly, ['ER', 'USL', 'LSL', 'UCL', 'LCL'], fontsize=11, loc='upper right') 
    lines_er_poly = ax_repl1a_er_poly.plot(df_repl1a_er_poly["Date (MM/DD/YYYY)"], df_repl1a_er_poly["Etch Rate (A/Min)"], visible=False)
    datacursor(lines_er_poly, hover=True, point_labels=df_repl1a_er_poly['Remarks'])
    # plt.show()        # shows 2 figures in different windows
    sht_repl1a_plot_er_poly.pictures.add(fig_repl1a_er_poly, name= "REPL1A_POLY_ER_Plot", update= True)

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw POLY Unif PLot
    fig_repl1a_unif_poly, ax_repl1a_unif_poly = plt.subplots(1,1, figsize=(20,6))
    monthyearFmt_unif_poly = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    ax_repl1a_unif_poly.xaxis.set_major_formatter(monthyearFmt_unif_poly)
    _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # ax_repl1a_unif_poly.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    ax_repl1a_unif_poly.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    ax_repl1a_unif_poly.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    plt.plot(df_repl1a_er_poly["Date (MM/DD/YYYY)"], df_repl1a_er_poly["% Uniformity"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    plt.plot(df_repl1a_er_poly["Date (MM/DD/YYYY)"], df_repl1a_er_poly["% Uni USL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    plt.plot(df_repl1a_er_poly["Date (MM/DD/YYYY)"], df_repl1a_er_poly["% Uni UCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    plt.xlabel('Date', fontsize=18)      # xlabel
    plt.ylabel('Poly Unif (%)', fontsize=18)     # ylabel
    # Custom Legends
    custom_lines_unif_poly = [
        Line2D([0], [0], color='#FF7F50', lw=4),	# ER
        Line2D([0], [0], color='#0000CD', lw=4),	# USL
        Line2D([0], [0], color='#FF1493', lw=4),	# UCL
        ]
    ax_repl1a_unif_poly.legend(custom_lines_unif_poly, ['ER', 'USL', 'LSL', 'UCL', 'LCL'], fontsize=11, loc='upper right') 
    lines_unif_poly = ax_repl1a_unif_poly.plot(df_repl1a_er_poly["Date (MM/DD/YYYY)"], df_repl1a_er_poly["% Uniformity"], visible=False)
    datacursor(lines_unif_poly, hover=True, point_labels=df_repl1a_er_poly['Remarks'])
    # plt.show()        # shows 2 figures in different windows
    sht_repl1a_plot_er_poly.pictures.add(fig_repl1a_unif_poly, name= "REPL1A_POLY_UNIF_Plot", update= True)


#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
# User Defined Functions (UDFs)
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
@xw.func
def hello(name):
    return "hello {0}".format(name)



