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


# def draw_plotly_reml1a_cp_plot(x, y0, y1, y2, remarks):
def draw_plotly_reml1a_cp_plot(x, y0, y1, remarks):
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
            title = 'CP Plot for REML1A',
            xaxis = dict(title= 'Date'),
            yaxis = dict(title= 'delta CP (no.s)')
        )
    fig = dict(data= data, layout = layout)
    py.offline.plot(fig, filename='REML1A_CP-Plot.html')
#------------------------------------------------------------------------------------------------------------------------------------------------------------------

# def draw_plotly_reml1c_cp_plot(x, y0, y1, y2, remarks):
def draw_plotly_reml1c_cp_plot(x, y0, y1, remarks):
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
            title = 'CP Plot for REML1C',
            xaxis = dict(title= 'Date'),
            yaxis = dict(title= 'delta CP (no.s)')
        )
    fig = dict(data= data, layout = layout)
    py.offline.plot(fig, filename='REML1C_CP-Plot.html')

#====================================================================================================================================================================
#####################################################################################################################################################################
def main():
    wb = xw.Book.caller()
    # wb.sheets[0].range("A1").value = "Hello xlwings!"		# test code

    #****************************************************************************************************************************************************************
    # Define sheets
    sht_reml1_cp = wb.sheets['REML1-CP']
    sht_reml1_er_ch_a_pr = wb.sheets['PR Ch A ER']
    sht_reml1_er_ch_c_pr = wb.sheets['PR Ch C ER']
    sht_reml1_plot_cp = wb.sheets['CP Plot']
    sht_reml1_plot_er_ch_a_pr = wb.sheets['PR Ch A Plot']
    sht_reml1_plot_er_ch_c_pr = wb.sheets['PR Ch C Plot']

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for CP
    df_reml1_cp = sht_reml1_cp.range('A9').options(
        pd.DataFrame, header=1, index=False, expand='table'
        ).value											                # fetch the data from sheet- 'ASBE1-CP'
    df_reml1_cp['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_reml1_cp = df_reml1_cp[["Date (MM/DD/YYYY)", "Chamber", "delta CP", "USL", "Remarks"]]        # The final dataframe with required columns
    # sht_reml1_plot_cp.range('A46').options(index=False).value = df_reml1_cp   	    # show the dataframe values into sheet- 'CP Plot'
    df_reml1_cp_ch_a = df_reml1_cp[df_reml1_cp["Chamber"] == 'DPS']					# dataframe for DPS (chamber A)
    # sht_reml1_plot_cp.range('A46').options(index=False).value = df_reml1_cp_ch_a           # show the dataframe values into sheet- 'CP Plot'
    df_reml1_cp_ch_c = df_reml1_cp[df_reml1_cp["Chamber"] == 'ASP']					# dataframe for ASP (chamber C)
    # sht_reml1_plot_cp.range('G46').options(index=False).value = df_reml1_cp_ch_c           # show the dataframe values into sheet- 'CP Plot'
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_reml1_cp_ch_a_date = df_reml1_cp_ch_a["Date (MM/DD/YYYY)"]
    df_reml1_cp_ch_a_delta_cp = df_reml1_cp_ch_a["delta CP"]
    df_reml1_cp_ch_a_usl = df_reml1_cp_ch_a["USL"]
    # df_reml1_cp_ch_a_ucl = df_reml1_cp_ch_a["UCL"]
    df_reml1_cp_ch_a_remarks = df_reml1_cp_ch_a["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_reml1_cp_ch_c_date = df_reml1_cp_ch_c["Date (MM/DD/YYYY)"]
    df_reml1_cp_ch_c_delta_cp = df_reml1_cp_ch_c["delta CP"]
    df_reml1_cp_ch_c_usl = df_reml1_cp_ch_c["USL"]
    # df_reml1_cp_ch_c_ucl = df_reml1_cp_ch_c["UCL"]
    df_reml1_cp_ch_c_remarks = df_reml1_cp_ch_c["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw CP Plot for DPS chamber i.e. Ch-A
    fig_reml1_cp_ch_a, ax_reml1_cp_ch_a = plt.subplots(1,1, figsize=(20,6))
    monthyearFmt_cp_ch_a = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    ax_reml1_cp_ch_a.xaxis.set_major_formatter(monthyearFmt_cp_ch_a)
    _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # ax_reml1_cp_ch_a.xaxis.set_major_locator(mdates.MonthLocator())          # set ticks after every Month
    ax_reml1_cp_ch_a.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    ax_reml1_cp_ch_a.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    plt.plot(df_reml1_cp_ch_a["Date (MM/DD/YYYY)"], df_reml1_cp_ch_a["delta CP"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs CP
    plt.plot(df_reml1_cp_ch_a["Date (MM/DD/YYYY)"], df_reml1_cp_ch_a["USL"], linestyle='-', color='#0000CD')        # plot date vs USL
    plt.xlabel('Date', fontsize=18)      # xlabel
    plt.ylabel('DPS delta CP (no.s)', fontsize=18)     # ylabel
    # Custom Legends
    custom_lines_cp_ch_a = [
        Line2D([0], [0], color='#FF7F50', lw=4),
        Line2D([0], [0], color='#0000CD', lw=4),
        ]
    ax_reml1_cp_ch_a.legend(custom_lines_cp_ch_a, ['CP', 'USL'], fontsize=11, loc='upper right')  
    lines_cp_ch_a = ax_reml1_cp_ch_a.plot(df_reml1_cp_ch_a["Date (MM/DD/YYYY)"], df_reml1_cp_ch_a["delta CP"], visible=False)
    # datacursor(lines_cp_ch_a, hover=True, point_labels=df_reml1_cp_ch_a['Remarks'])
    # plt.show()
    # sht_cp_plot.activate()
    # pic_cp = plt.show()
    # plt.show('ASBE1_CP_Plot', left=xw.Range('A1').left, top=xw.Range('A1').top)      # this would activate hover 
    # sh_plot_cp.pictures.add(pic_cp, name= "ASFE1_CP_Plot", update= True)
    sht_reml1_plot_cp.pictures.add(fig_reml1_cp_ch_a, name= "REML1_DPS_CP_Plot", update= True)


    # Draw CP Plot (using Plotly) in Browser 
    draw_plotly_reml1a_cp_plot(
        x = df_reml1_cp_ch_a_date, 
        y0 = df_reml1_cp_ch_a_delta_cp, 
        y1 = df_reml1_cp_ch_a_usl, 
        # y2 = df_repl1a_cp_ucl,
        remarks = df_reml1_cp_ch_a_remarks
        )

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw CP Plot for ASP chamber i.e. Ch-C
    fig_reml1_cp_ch_c, ax_reml1_cp_ch_c = plt.subplots(1,1, figsize=(20,6))
    monthyearFmt_cp_ch_c = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    ax_reml1_cp_ch_c.xaxis.set_major_formatter(monthyearFmt_cp_ch_c)
    _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # ax_reml1_cp_ch_c.xaxis.set_major_locator(mdates.MonthLocator())          # set ticks after every Month
    ax_reml1_cp_ch_c.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    ax_reml1_cp_ch_c.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    plt.plot(df_reml1_cp_ch_c["Date (MM/DD/YYYY)"], df_reml1_cp_ch_c["delta CP"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs CP
    plt.plot(df_reml1_cp_ch_c["Date (MM/DD/YYYY)"], df_reml1_cp_ch_c["USL"], linestyle='-', color='#0000CD')        # plot date vs USL
    plt.xlabel('Date', fontsize=18)      # xlabel
    plt.ylabel('ASP delta CP (no.s)', fontsize=18)     # ylabel
    # Custom Legends
    custom_lines_cp_ch_c = [
        Line2D([0], [0], color='#FF7F50', lw=4),
        Line2D([0], [0], color='#0000CD', lw=4),
        ]
    ax_reml1_cp_ch_c.legend(custom_lines_cp_ch_c, ['CP', 'USL'], fontsize=11, loc='upper right')  
    lines_cp_ch_c = ax_reml1_cp_ch_c.plot(df_reml1_cp_ch_c["Date (MM/DD/YYYY)"], df_reml1_cp_ch_c["delta CP"], visible=False)
    # datacursor(lines_cp_ch_c, hover=True, point_labels=df_reml1_cp_ch_c['Remarks'])
    # plt.show()
    # sht_cp_plot.activate()
    # pic_cp = plt.show()
    # plt.show('ASBE1_CP_Plot', left=xw.Range('A1').left, top=xw.Range('A1').top)      # this would activate hover 
    # sht_reml1_plot_cp.pictures.add(pic_cp, name= "ASFE1_CP_Plot", update= True)
    sht_reml1_plot_cp.pictures.add(fig_reml1_cp_ch_c, name= "REML1_ASP_CP_Plot", update= True)

    # Draw CP Plot (using Plotly) in Browser 
    draw_plotly_reml1c_cp_plot(
        x = df_reml1_cp_ch_c_date, 
        y0 = df_reml1_cp_ch_c_delta_cp, 
        y1 = df_reml1_cp_ch_c_usl, 
        # y2 = df_repl1a_cp_ucl,
        remarks = df_reml1_cp_ch_c_remarks
        )

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for Ch A PR i.e. DPS chamber 
    # data_folder = Path(os.getcwd())
    # file_to_open = data_folder / "ASH09_QC_LOG_BOOK.xlsm"
    # excel_file = pd.ExcelFile(file_to_open)

    excel_file_sht_ch_a_pr = pd.ExcelFile("H:\\excel\\dryetch\\Excel-office\\macro_enabled_logbooks\\CNT02_QC_LOG_BOOK\\CNT02_QC_LOG_BOOK.xlsm")
    # excel_file_sht_ch_a_pr = pd.ExcelFile("\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\CNT_02_LOG_BOOK\\CNT02_QC_LOG_BOOK_macro\\CNT02_QC_LOG_BOOK.xlsm")
    df_reml1_er_ch_a_pr = excel_file_sht_ch_a_pr.parse('PR Ch A ER', skiprows=8)                            # copy a sheet and paste into another sheet and skiprows 8
    
    df_reml1_er_ch_a_pr = df_reml1_er_ch_a_pr[["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uniformity", "LSL", "USL", "LCL", "UCL", "Remarks", "% Uni USL", "% Uni UCL"]]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_reml1_er_ch_a_pr['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_reml1_er_ch_a_pr = df_reml1_er_ch_a_pr.dropna()                                              # dropping rows where at least one element is missing
    # sht_plot_nit.range('A28').options(index=False).value = df_reml1_er_ch_a_pr        # show the dataframe values into sheet- 'CP Plot'
    
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw Ch A PR ER PLot
    fig_reml1_er_ch_a_pr, ax_reml1_er_ch_a_pr = plt.subplots(1,1, figsize=(20,6))
    monthyearFmt_er_pr = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    ax_reml1_er_ch_a_pr.xaxis.set_major_formatter(monthyearFmt_er_pr)
    _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # ax_reml1_er_ch_a_pr.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    ax_reml1_er_ch_a_pr.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    ax_reml1_er_ch_a_pr.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    plt.plot(df_reml1_er_ch_a_pr["Date (MM/DD/YYYY)"], df_reml1_er_ch_a_pr["Etch Rate (A/Min)"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    plt.plot(df_reml1_er_ch_a_pr["Date (MM/DD/YYYY)"], df_reml1_er_ch_a_pr["USL"], linestyle='-', color='#0000CD')        # plot date vs USL
    plt.plot(df_reml1_er_ch_a_pr["Date (MM/DD/YYYY)"], df_reml1_er_ch_a_pr["LSL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    plt.plot(df_reml1_er_ch_a_pr["Date (MM/DD/YYYY)"], df_reml1_er_ch_a_pr["UCL"], linestyle='-', color='#FF1493')        # plot date vs UCL
    plt.plot(df_reml1_er_ch_a_pr["Date (MM/DD/YYYY)"], df_reml1_er_ch_a_pr["LCL"], linestyle='-', color='#FF1493')        # plot date vs LCL
    plt.xlabel('Date', fontsize=18)      # xlabel
    plt.ylabel('Ch A PR ER (A/min)', fontsize=18)     # ylabel
    # Custom Legends
    custom_lines_er_cha_pr = [
        Line2D([0], [0], color='#FF7F50', lw=4),	# ER
        Line2D([0], [0], color='#0000CD', lw=4),	# USL
        Line2D([0], [0], color='#0000CD', lw=4),	# LSL
        Line2D([0], [0], color='#FF1493', lw=4),	# UCL
        Line2D([0], [0], color='#FF1493', lw=4)		# LCL        
        ]
    ax_reml1_er_ch_a_pr.legend(custom_lines_er_cha_pr, ['ER', 'USL', 'LSL', 'UCL', 'LCL'], fontsize=11, loc='upper right') 
    lines_er_ch_a_pr = ax_reml1_er_ch_a_pr.plot(df_reml1_er_ch_a_pr["Date (MM/DD/YYYY)"], df_reml1_er_ch_a_pr["Etch Rate (A/Min)"], visible=False)
    # datacursor(lines_er_ch_a_pr, hover=True, point_labels=df_reml1_er_ch_a_pr['Remarks'])
    # plt.show()        # shows 2 figures in different windows
    sht_reml1_plot_er_ch_a_pr.pictures.add(fig_reml1_er_ch_a_pr, name= "REML1_CH_A_PR_ER_Plot", update= True)

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
	# Draw ChA PR Unif PLot
    fig_reml1_unif_ch_a_pr, ax_reml1_unif_ch_a_pr = plt.subplots(1,1, figsize=(20,6))
    monthyearFmt_unif_ch_a_pr = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    ax_reml1_unif_ch_a_pr.xaxis.set_major_formatter(monthyearFmt_unif_ch_a_pr)
    _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # ax_reml1_unif_ch_a_pr.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    ax_reml1_unif_ch_a_pr.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    ax_reml1_unif_ch_a_pr.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    plt.plot(df_reml1_er_ch_a_pr["Date (MM/DD/YYYY)"], df_reml1_er_ch_a_pr["% Uniformity"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    plt.plot(df_reml1_er_ch_a_pr["Date (MM/DD/YYYY)"], df_reml1_er_ch_a_pr["% Uni USL"], linestyle='-', color='#0000CD')        # plot date vs USL
    plt.plot(df_reml1_er_ch_a_pr["Date (MM/DD/YYYY)"], df_reml1_er_ch_a_pr["% Uni UCL"], linestyle='-', color='#FF1493')        # plot date vs UCL
    plt.xlabel('Date', fontsize=18)      # xlabel
    plt.ylabel('Ch A PR Unif (%)', fontsize=18)     # ylabel
    # Custom Legends
    custom_lines_unif_ch_a_pr = [
        Line2D([0], [0], color='#FF7F50', lw=4),	# ER
        Line2D([0], [0], color='#0000CD', lw=4),	# USL
        Line2D([0], [0], color='#FF1493', lw=4),	# UCL
        ]
    ax_reml1_unif_ch_a_pr.legend(custom_lines_unif_ch_a_pr, ['Unif', 'USL', 'UCL'], fontsize=11, loc='upper right') 
    lines_unif_ch_a_pr = ax_reml1_unif_ch_a_pr.plot(df_reml1_er_ch_a_pr["Date (MM/DD/YYYY)"], df_reml1_er_ch_a_pr["% Uniformity"], visible=False)
    # datacursor(lines_unif_ch_a_pr, hover=True, point_labels=df_reml1_er_ch_a_pr['Remarks'])
    # plt.show()        # shows 2 figures in different windows
    sht_reml1_plot_er_ch_a_pr.pictures.add(fig_reml1_unif_ch_a_pr, name= "REML1_CH_A_PR_UNIF_Plot", update= True)


    #****************************************************************************************************************************************************************
    # Fetch Dataframe for Ch C PR i.e. ASP Chamber
    # data_folder = Path(os.getcwd())
    # file_to_open = data_folder / "ASH09_QC_LOG_BOOK.xlsm"
    # excel_file = pd.ExcelFile(file_to_open)

    excel_file_sht_ch_c_pr = pd.ExcelFile("H:\\excel\\dryetch\\Excel-office\\macro_enabled_logbooks\\CNT02_QC_LOG_BOOK\\CNT02_QC_LOG_BOOK.xlsm")
    # excel_file_sht_ch_c_pr = pd.ExcelFile("\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\CNT_02_LOG_BOOK\\CNT02_QC_LOG_BOOK_macro\\CNT02_QC_LOG_BOOK.xlsm")
    df_reml1_er_ch_c_pr = excel_file_sht_ch_c_pr.parse('PR Ch C ER', skiprows=8)                            # copy a sheet and paste into another sheet and skiprows 8
    
    df_reml1_er_ch_c_pr = df_reml1_er_ch_c_pr[["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uniformity", "LSL", "USL", "LCL", "UCL", "Remarks", "% Uni USL"]]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_reml1_er_ch_c_pr['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_reml1_er_ch_c_pr = df_reml1_er_ch_c_pr.dropna()                                              # dropping rows where at least one element is missing
    # sht_reml1_plot_er_ch_c_pr.range('A28').options(index=False).value = df_reml1_er_ch_c_pr        # show the dataframe values into sheet- 'CP Plot'
    
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw ChC ER PLot
    fig_reml1_er_ch_c_pr, ax_reml1_er_ch_c_pr = plt.subplots(1,1, figsize=(20,6))
    monthyearFmt_er_ch_c_pr = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    ax_reml1_er_ch_c_pr.xaxis.set_major_formatter(monthyearFmt_er_ch_c_pr)
    _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # ax_reml1_er_ch_c_pr.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    ax_reml1_er_ch_c_pr.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    ax_reml1_er_ch_c_pr.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    plt.plot(df_reml1_er_ch_c_pr["Date (MM/DD/YYYY)"], df_reml1_er_ch_c_pr["Etch Rate (A/Min)"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    plt.plot(df_reml1_er_ch_c_pr["Date (MM/DD/YYYY)"], df_reml1_er_ch_c_pr["USL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    plt.plot(df_reml1_er_ch_c_pr["Date (MM/DD/YYYY)"], df_reml1_er_ch_c_pr["LSL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    plt.plot(df_reml1_er_ch_c_pr["Date (MM/DD/YYYY)"], df_reml1_er_ch_c_pr["UCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    plt.plot(df_reml1_er_ch_c_pr["Date (MM/DD/YYYY)"], df_reml1_er_ch_c_pr["LCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    plt.xlabel('Date', fontsize=18)      # xlabel
    plt.ylabel('Ch C PR ER (A/min)', fontsize=18)     # ylabel
    # Custom Legends
    custom_lines_er_ch_c_pr = [
        Line2D([0], [0], color='#FF7F50', lw=4),	# ER
        Line2D([0], [0], color='#0000CD', lw=4),	# USL
        Line2D([0], [0], color='#0000CD', lw=4),	# LSL
        Line2D([0], [0], color='#FF1493', lw=4),	# UCL
        Line2D([0], [0], color='#FF1493', lw=4)		# LCL        
        ]
    ax_reml1_er_ch_c_pr.legend(custom_lines_er_ch_c_pr, ['ER', 'USL', 'LSL', 'UCL', 'LCL'], fontsize=11, loc='upper right') 
    lines_er_ch_c_pr = ax_reml1_er_ch_c_pr.plot(df_reml1_er_ch_c_pr["Date (MM/DD/YYYY)"], df_reml1_er_ch_c_pr["Etch Rate (A/Min)"], visible=False)
    # datacursor(lines_er_ch_c_pr, hover=True, point_labels=df_reml1_er_ch_c_pr['Remarks'])
    # plt.show()        # shows 2 figures in different windows
    sht_reml1_plot_er_ch_c_pr.pictures.add(fig_reml1_er_ch_c_pr, name= "REML1_CH_C_PR_ER_Plot", update= True)

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw ChC UNIF PLot
    fig_reml1_unif_ch_c_pr, ax_reml1_unif_ch_c_pr = plt.subplots(1,1, figsize=(20,6))
    monthyearFmt_unif_ch_c_pr = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    ax_reml1_unif_ch_c_pr.xaxis.set_major_formatter(monthyearFmt_unif_ch_c_pr)
    _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # ax_reml1_unif_ch_c_pr.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    ax_reml1_unif_ch_c_pr.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    ax_reml1_unif_ch_c_pr.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    plt.plot(df_reml1_er_ch_c_pr["Date (MM/DD/YYYY)"], df_reml1_er_ch_c_pr["% Uniformity"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    plt.plot(df_reml1_er_ch_c_pr["Date (MM/DD/YYYY)"], df_reml1_er_ch_c_pr["% Uni USL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    plt.xlabel('Date', fontsize=18)      # xlabel
    plt.ylabel('Ch C PR Unif (%)', fontsize=18)     # ylabel
    # Custom Legends
    custom_lines_unif_ch_c_pr = [
        Line2D([0], [0], color='#FF7F50', lw=4),	# ER
        Line2D([0], [0], color='#0000CD', lw=4),	# USL
        Line2D([0], [0], color='#FF1493', lw=4),	# UCL
        ]
    ax_reml1_unif_ch_c_pr.legend(custom_lines_unif_ch_c_pr, ['Unif', 'USL'], fontsize=11, loc='upper right') 
    lines_unif_ch_c_pr = ax_reml1_unif_ch_c_pr.plot(df_reml1_er_ch_c_pr["Date (MM/DD/YYYY)"], df_reml1_er_ch_c_pr["% Uniformity"], visible=False)
    # datacursor(lines_unif_ch_c_pr, hover=True, point_labels=df_reml1_er_ch_c_pr['Remarks'])
    # plt.show()        # shows 2 figures in different windows
    sht_reml1_plot_er_ch_c_pr.pictures.add(fig_reml1_unif_ch_c_pr, name= "REML1_CH_C_PR_UNIF_Plot", update= True)


#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
# User Defined Functions (UDFs)
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
@xw.func
def hello(name):
    return "hello {0}".format(name)



