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
cp_chart_title = 'CP Plot for REPL1A'  # title for CP chart
cp_chart_xlabel = 'Date'   # xaxis name for CP chart
cp_chart_ylabel = 'delta CP (no.s)'     # yaxis name for CP chart
cp_chart_html_file = 'REPL1A_CP-Plot.html'   # HTML filename for CP chart
cp_chart_trace_count = 2    # no. of traces in CP chart
er_nit_chart_title = 'Nit ER Plot for REPL1A'  # title for SiN ER chart
er_nit_chart_xlabel = 'Date'        # xaxis name for SiN ER chart
er_nit_chart_ylabel = 'Nit ER (A/min)'   # yaxis name for SiN ER chart
er_nit_chart_html_file = 'REPL1A_Nit_ER-Plot.html'   # HTML filename for SiN ER chart
er_nit_chart_trace_count = 5    # no. of traces in SiN ER chart
unif_nit_chart_title = 'Nit Uniformity Plot for REPL1A'  # title for SiN Unif chart
unif_nit_chart_xlabel = 'Date'      # xaxis name for SiN Unif chart
unif_nit_chart_ylabel = 'Nit Unif (%)'    # yaxis name for SiN Unif chart
unif_nit_chart_html_file = 'REPL1A_Nit_Unif-Plot.html'   # HTML filename for SiN Unif chart
unif_nit_chart_trace_count = 3    # no. of traces in SiN Unif chart
er_poly_chart_title = 'Poly ER Plot for REPL1A'  # title for Poly ER chart
er_poly_chart_xlabel = 'Date'        # xaxis name for Poly ER chart
er_poly_chart_ylabel = 'Poly ER (A/min)'   # yaxis name for Poly ER chart
er_poly_chart_html_file = 'REPL1A_Poly_ER-Plot.html'   # HTML filename for Poly ER chart
er_poly_chart_trace_count = 5    # no. of traces in Poly ER chart
unif_poly_chart_title = 'Poly Uniformity Plot for REPL1A'  # title for Poly Unif chart
unif_poly_chart_xlabel = 'Date'      # xaxis name for Poly Unif chart
unif_poly_chart_ylabel = 'Poly Unif (%)'    # yaxis name for Poly Unif chart
unif_poly_chart_html_file = 'REPL1A_Poly_Unif-Plot.html'   # HTML filename for Poly Unif chart
unif_poly_chart_trace_count = 3    # no. of traces in Poly Unif chart

excel_file_directory = "I:\\excel\\dryetch\\Excel-office\\macro_enabled_logbooks\\CNT01_QC_LOG_BOOK_Ch_A\\CNT01_QC_LOG_BOOK_Ch_A.xlsm"
# excel_file_directory = "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\CNT_01_LOG_BOOK\\CNT01_QC_LOG_BOOK_Ch_A_macro\\CNT01_QC_LOG_BOOK_Ch_A.xlsm"

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots CP Chart with `cp_chart_trace_count` traces v/s Date.
"x": Date (x-axis) for CP Chart
"y1": Delta-CP (y-axis) for CP Chart
"y2": USL (y-axis) for CP Chart
# "y3": UCL (y-axis) for CP Chart
"""
def draw_plotly_repl1a_cp_plot(x, y1, y2, remarks):
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

    # data = [trace1, trace2, trace3]
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
"Description": This function plots ER Chart with `er_nit_chart_trace_count` traces v/s Date.
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
            name = 'SiN ER',
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
            title = er_nit_chart_title,
            xaxis = dict(title= er_nit_chart_xlabel),
            yaxis = dict(title= er_nit_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= er_nit_chart_html_file)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with `unif_nit_chart_trace_count` traces v/s Date.
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
            title = unif_nit_chart_title,
            xaxis = dict(title= unif_nit_chart_xlabel),
            yaxis = dict(title= unif_nit_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= unif_nit_chart_html_file)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots ER Chart with `er_poly_chart_trace_count` traces v/s Date.
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
            name = 'Poly ER',
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
            title = er_poly_chart_title,
            xaxis = dict(title= er_poly_chart_xlabel),
            yaxis = dict(title= er_poly_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= er_poly_chart_html_file)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with `unif_nit_chart_trace_count` traces v/s Date.
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
            title = unif_poly_chart_title,
            xaxis = dict(title= unif_poly_chart_xlabel),
            yaxis = dict(title= unif_poly_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= unif_poly_chart_html_file)



#====================================================================================================================================================================
#####################################################################################################################################################################
def main():
    wb = xw.Book.caller()
    # wb.sheets[0].range("A1").value = "Hello xlwings!"     # test code

    #****************************************************************************************************************************************************************
    # Define sheets
    sht_repl1a_cp = wb.sheets['REPL1A-CP']
    sht_repl1a_er_nit = wb.sheets['REPL1A-ERNit']
    sht_repl1a_er_poly = wb.sheets['REPL1A-ERPoly']
    # sht_repl1a_plot_cp = wb.sheets['CP Plot']
    # sht_repl1a_plot_er_nit = wb.sheets['SiN Plot']
    # sht_repl1a_plot_er_poly = wb.sheets['Poly Plot']

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for CP Plot
    df_repl1a_cp = sht_repl1a_cp.range('A9').options(
        pd.DataFrame, header=1, index=False, expand='table'
        ).value                                                         # fetch the data from sheet- 'ASBE1-CP'
    df_repl1a_cp['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_repl1a_cp = df_repl1a_cp[["Date (MM/DD/YYYY)", "delta CP", "USL", "Remarks"]]        # The final dataframe with required columns
    df_repl1a_cp = df_repl1a_cp.dropna()                                              # dropping rows where at least one element is misnitg
    # sht_repl1a_plot_cp.range('A25').options(index=False).value = df_repl1a_cp         # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_repl1a_cp_date = df_repl1a_cp["Date (MM/DD/YYYY)"]
    df_repl1a_cp_delta_cp = df_repl1a_cp["delta CP"]
    df_repl1a_cp_usl = df_repl1a_cp["USL"]
    # df_repl1a_cp_ucl = df_repl1a_cp["UCL"]
    df_repl1a_cp_remarks = df_repl1a_cp["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # # Draw CP Plot
    # fig_repl1a_cp, fig_repl1a_cp = plt.subplots(1,1, figsize=(20,6))
    # monthyearFmt_cp = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    # fig_repl1a_cp.xaxis.set_major_formatter(monthyearFmt_cp)
    # _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # # fig_repl1a_cp.xaxis.set_major_locator(mdates.MonthLocator())          # set ticks after every Month
    # fig_repl1a_cp.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    # fig_repl1a_cp.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    # plt.plot(df_repl1a_cp["Date (MM/DD/YYYY)"], df_repl1a_cp["delta CP"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs CP
    # plt.plot(df_repl1a_cp["Date (MM/DD/YYYY)"], df_repl1a_cp["USL"], linestyle='-', color='#0000CD')        # plot date vs USL
    # plt.xlabel('Date', fontsize=18)      # xlabel
    # plt.ylabel('delta CP (no.s)', fontsize=18)     # ylabel
    # # Custom Legends
    # custom_lines_cp = [
    #     Line2D([0], [0], color='#FF7F50', lw=4),
    #     Line2D([0], [0], color='#0000CD', lw=4),
    #     ]
    # fig_repl1a_cp.legend(custom_lines_cp, ['CP', 'USL'], fontsize=11, loc='upper right')  
    # lines_cp = fig_repl1a_cp.plot(df_repl1a_cp["Date (MM/DD/YYYY)"], df_repl1a_cp["delta CP"], visible=False)
    # # datacursor(lines_cp, hover=True, point_labels=df_repl1a_cp['Remarks'])
    # # plt.show()
    # # sht_repl1a_plot_cp.activate()
    # # pic_cp = plt.show()
    # # plt.show('ASBE1_CP_Plot', left=xw.Range('A1').left, top=xw.Range('A1').top)      # this would activate hover 
    # # sht_repl1a_plot_cp.pictures.add(pic_cp, name= "REPL1A_CP_Plot", update= True)
    # # sht_repl1a_plot_cp.pictures.add(fig_repl1a_cp, name= "REPL1A_CP_Plot", update= True)

    # Draw CP Plot (unitg Plotly) in Browser 
    draw_plotly_repl1a_cp_plot(
        x = df_repl1a_cp_date, 
        y1 = df_repl1a_cp_delta_cp, 
        y2 = df_repl1a_cp_usl, 
        # y3 = df_repl1a_cp_ucl,
        remarks = df_repl1a_cp_remarks
        )

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for NIT ER & UNif PLot
    # data_folder = Path(os.getcwd())
    # file_to_open = data_folder / "ASH09_QC_LOG_BOOK.xlsm"
    # excel_file = pd.ExcelFile(file_to_open)

    excel_file_sht_nit = pd.ExcelFile(excel_file_directory)
    df_repl1a_er_nit = excel_file_sht_nit.parse('REPL1A-ERNit', skiprows=9)                            # copy a sheet and paste into another sheet and skiprows 9
    
    df_repl1a_er_nit = df_repl1a_er_nit[["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "LSL", "USL", "LCL", "UCL", "Remarks", "% Uni USL", "% Uni UCL"]]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_repl1a_er_nit['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_repl1a_er_nit = df_repl1a_er_nit.dropna()                                              # dropping rows where at least one element is misnitg
    # sht_repl1a_plot_er_nit.range('A28').options(index=False).value = df_repl1a_er_nit        # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param for SiN ER & Unif PLot
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
    # # Draw NIT ER PLot
    # fig_repl1a_er_nit, ax_repl1a_er_nit = plt.subplots(1,1, figsize=(20,6))
    # monthyearFmt_er_nit = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    # ax_repl1a_er_nit.xaxis.set_major_formatter(monthyearFmt_er_nit)
    # _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # # ax_repl1a_er_nit.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    # ax_repl1a_er_nit.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    # ax_repl1a_er_nit.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    # plt.plot(df_repl1a_er_nit["Date (MM/DD/YYYY)"], df_repl1a_er_nit["Etch Rate (A/Min)"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    # plt.plot(df_repl1a_er_nit["Date (MM/DD/YYYY)"], df_repl1a_er_nit["USL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # plt.plot(df_repl1a_er_nit["Date (MM/DD/YYYY)"], df_repl1a_er_nit["LSL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # plt.plot(df_repl1a_er_nit["Date (MM/DD/YYYY)"], df_repl1a_er_nit["UCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    # plt.plot(df_repl1a_er_nit["Date (MM/DD/YYYY)"], df_repl1a_er_nit["LCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    # plt.xlabel('Date', fontsize=18)      # xlabel
    # plt.ylabel('SiN ER (A/min)', fontsize=18)     # ylabel
    # # Custom Legends
    # custom_lines_er_nit = [
    #     Line2D([0], [0], color='#FF7F50', lw=4),  # ER
    #     Line2D([0], [0], color='#0000CD', lw=4),  # USL
    #     Line2D([0], [0], color='#0000CD', lw=4),  # LSL
    #     Line2D([0], [0], color='#FF1493', lw=4),  # UCL
    #     Line2D([0], [0], color='#FF1493', lw=4)       # LCL        
    #     ]
    # ax_repl1a_er_nit.legend(custom_lines_er_nit, ['ER', 'USL', 'LSL', 'UCL', 'LCL'], fontsize=11, loc='upper right') 
    # lines_er_nit = ax_repl1a_er_nit.plot(df_repl1a_er_nit["Date (MM/DD/YYYY)"], df_repl1a_er_nit["Etch Rate (A/Min)"], visible=False)
    # # datacursor(lines_er_nit, hover=True, point_labels=df_repl1a_er_nit['Remarks'])
    # # plt.show()        # shows 2 figures in different windows
    # sht_repl1a_plot_er_nit.pictures.add(fig_repl1a_er_nit, name= "REPL1A_NIT_ER_Plot", update= True)

    # Draw Nit ER Plot (unitg Plotly) in Browser 
    draw_plotly_repl1a_er_nit_plot(
        x = df_repl1a_er_nit_date, 
        y1 = df_repl1a_er_nit_er,
        y2 = df_repl1a_er_nit_usl, 
        y3 = df_repl1a_er_nit_lsl,
        y4 = df_repl1a_er_nit_ucl,
        y5 = df_repl1a_er_nit_lcl,
        remarks = df_repl1a_er_nit_remarks
        )
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw Nit Unif Plot (unitg Plotly) in Browser     
    draw_plotly_repl1a_unif_nit_plot(
        x = df_repl1a_er_nit_date, 
        y1 = df_repl1a_er_nit_unif, 
        y2 = df_repl1a_er_nit_unif_usl,
        y3 = df_repl1a_er_nit_unif_ucl,
        remarks = df_repl1a_er_nit_remarks
        )
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # # Draw NIT Unif PLot
 #    fig_repl1a_unif_nit, ax_repl1a_unif_nit = plt.subplots(1,1, figsize=(20,6))
 #    monthyearFmt_unif_nit = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
 #    ax_repl1a_unif_nit.xaxis.set_major_formatter(monthyearFmt_unif_nit)
 #    _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
 #    # ax_repl1a_unif_nit.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
 #    ax_repl1a_unif_nit.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
 #    ax_repl1a_unif_nit.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
 #    plt.plot(df_repl1a_er_nit["Date (MM/DD/YYYY)"], df_repl1a_er_nit["% Uniformity"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
 #    plt.plot(df_repl1a_er_nit["Date (MM/DD/YYYY)"], df_repl1a_er_nit["% Uni USL"], linestyle='-', color='#0000CD')        # plot date vs LSL
 #    plt.plot(df_repl1a_er_nit["Date (MM/DD/YYYY)"], df_repl1a_er_nit["% Uni UCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
 #    plt.xlabel('Date', fontsize=18)      # xlabel
 #    plt.ylabel('SiN Unif (%)', fontsize=18)     # ylabel
 #    # Custom Legends
 #    custom_lines_unif_nit = [
 #        Line2D([0], [0], color='#FF7F50', lw=4),  # ER
 #        Line2D([0], [0], color='#0000CD', lw=4),  # USL
 #        Line2D([0], [0], color='#FF1493', lw=4),  # UCL
 #        ]
 #    ax_repl1a_unif_nit.legend(custom_lines_unif_nit, ['Unif', 'USL', 'UCL'], fontsize=11, loc='upper right') 
 #    lines_unif_nit = ax_repl1a_unif_nit.plot(df_repl1a_er_nit["Date (MM/DD/YYYY)"], df_repl1a_er_nit["% Uniformity"], visible=False)
 #    # datacursor(lines_unif_nit, hover=True, point_labels=df_repl1a_er_nit['Remarks'])
 #    # plt.show()        # shows 2 figures in different windows
 #    sht_repl1a_plot_er_nit.pictures.add(fig_repl1a_unif_nit, name= "REPL1A_NIT_UNIF_Plot", update= True)


    #****************************************************************************************************************************************************************
    # Fetch Dataframe for POLY ER & Unif PLot  
    # data_folder = Path(os.getcwd())
    # file_to_open = data_folder / "ASH09_QC_LOG_BOOK.xlsm"
    # excel_file = pd.ExcelFile(file_to_open)

    excel_file_sht_poly = pd.ExcelFile(excel_file_directory)
    df_repl1a_er_poly = excel_file_sht_poly.parse('REPL1A-ERPoly', skiprows=9)                            # copy a sheet and paste into another sheet and skiprows 9
    
    df_repl1a_er_poly = df_repl1a_er_poly[["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "LSL", "USL", "LCL", "UCL", "Remarks", "% Uni USL", "% Uni UCL"]]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_repl1a_er_poly['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_repl1a_er_poly = df_repl1a_er_poly.dropna()                                              # dropping rows where at least one element is misnitg
    # sht_repl1a_plot_er_poly.range('A28').options(index=False).value = df_repl1a_er_poly        # show the dataframe values into sheet- 'CP Plot'
 
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param for SiN ER & Unif PLot
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
    # # Draw POLY ER PLot
    # fig_repl1a_er_poly, ax_repl1a_er_poly = plt.subplots(1,1, figsize=(20,6))
    # monthyearFmt_er_poly = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    # ax_repl1a_er_poly.xaxis.set_major_formatter(monthyearFmt_er_poly)
    # _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # # ax_repl1a_er_poly.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    # ax_repl1a_er_poly.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    # ax_repl1a_er_poly.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    # plt.plot(df_repl1a_er_poly["Date (MM/DD/YYYY)"], df_repl1a_er_poly["Etch Rate (A/Min)"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    # plt.plot(df_repl1a_er_poly["Date (MM/DD/YYYY)"], df_repl1a_er_poly["USL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # plt.plot(df_repl1a_er_poly["Date (MM/DD/YYYY)"], df_repl1a_er_poly["LSL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # plt.plot(df_repl1a_er_poly["Date (MM/DD/YYYY)"], df_repl1a_er_poly["UCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    # plt.plot(df_repl1a_er_poly["Date (MM/DD/YYYY)"], df_repl1a_er_poly["LCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    # plt.xlabel('Date', fontsize=18)      # xlabel
    # plt.ylabel('Poly ER (A/min)', fontsize=18)     # ylabel
    # # Custom Legends
    # custom_lines_er_poly = [
    #     Line2D([0], [0], color='#FF7F50', lw=4),    # ER
    #     Line2D([0], [0], color='#0000CD', lw=4),    # USL
    #     Line2D([0], [0], color='#0000CD', lw=4),    # LSL
    #     Line2D([0], [0], color='#FF1493', lw=4),    # UCL
    #     Line2D([0], [0], color='#FF1493', lw=4)     # LCL        
    #     ]
    # ax_repl1a_er_poly.legend(custom_lines_er_poly, ['ER', 'USL', 'LSL', 'UCL', 'LCL'], fontsize=11, loc='upper right') 
    # lines_er_poly = ax_repl1a_er_poly.plot(df_repl1a_er_poly["Date (MM/DD/YYYY)"], df_repl1a_er_poly["Etch Rate (A/Min)"], visible=False)
    # # datacursor(lines_er_poly, hover=True, point_labels=df_repl1a_er_poly['Remarks'])
    # # plt.show()        # shows 2 figures in different windows
    # sht_repl1a_plot_er_poly.pictures.add(fig_repl1a_er_poly, name= "REPL1A_POLY_ER_Plot", update= True)

    # Draw Poly ER Plot (using Plotly) in Browser 
    draw_plotly_repl1a_er_poly_plot(
        x = df_repl1a_er_poly_date, 
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
        x = df_repl1a_er_poly_date, 
        y1 = df_repl1a_er_poly_unif, 
        y2 = df_repl1a_er_poly_unif_usl,
        y3 = df_repl1a_er_poly_unif_ucl,
        remarks = df_repl1a_er_poly_remarks
        )

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw POLY Unif PLot
    # fig_repl1a_unif_poly, ax_repl1a_unif_poly = plt.subplots(1,1, figsize=(20,6))
    # monthyearFmt_unif_poly = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    # ax_repl1a_unif_poly.xaxis.set_major_formatter(monthyearFmt_unif_poly)
    # _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # # ax_repl1a_unif_poly.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    # ax_repl1a_unif_poly.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    # ax_repl1a_unif_poly.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    # plt.plot(df_repl1a_er_poly["Date (MM/DD/YYYY)"], df_repl1a_er_poly["% Uniformity"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    # plt.plot(df_repl1a_er_poly["Date (MM/DD/YYYY)"], df_repl1a_er_poly["% Uni USL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # plt.plot(df_repl1a_er_poly["Date (MM/DD/YYYY)"], df_repl1a_er_poly["% Uni UCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    # plt.xlabel('Date', fontsize=18)      # xlabel
    # plt.ylabel('Poly Unif (%)', fontsize=18)     # ylabel
    # # Custom Legends
    # custom_lines_unif_poly = [
    #     Line2D([0], [0], color='#FF7F50', lw=4),    # ER
    #     Line2D([0], [0], color='#0000CD', lw=4),    # USL
    #     Line2D([0], [0], color='#FF1493', lw=4),    # UCL
    #     ]
    # ax_repl1a_unif_poly.legend(custom_lines_unif_poly, ['ER', 'USL', 'LSL', 'UCL', 'LCL'], fontsize=11, loc='upper right') 
    # lines_unif_poly = ax_repl1a_unif_poly.plot(df_repl1a_er_poly["Date (MM/DD/YYYY)"], df_repl1a_er_poly["% Uniformity"], visible=False)
    # # datacursor(lines_unif_poly, hover=True, point_labels=df_repl1a_er_poly['Remarks'])
    # # plt.show()        # shows 2 figures in different windows
    # sht_repl1a_plot_er_poly.pictures.add(fig_repl1a_unif_poly, name= "REPL1A_POLY_UNIF_Plot", update= True)


#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
# User Defined Functions (UDFs)
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
@xw.func
def hello(name):
    return "hello {0}".format(name)



