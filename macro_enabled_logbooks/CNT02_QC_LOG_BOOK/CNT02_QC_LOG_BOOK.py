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
cp_cha_chart_title = 'CP Plot for REML1A'  # title for CP chart
cp_cha_chart_xlabel = 'Date'   # xaxis name for CP chart
cp_cha_chart_ylabel = 'delta CP (no.s)'     # yaxis name for CP chart
cp_cha_chart_html_file = 'REML1A_CP-Plot.html'   # HTML filename for CP chart
cp_cha_chart_trace_count = 2    # no. of traces in CP chart
cp_chc_chart_title = 'CP Plot for REML1C'  # title for CP chart
cp_chc_chart_xlabel = 'Date'   # xaxis name for CP chart
cp_chc_chart_ylabel = 'delta CP (no.s)'     # yaxis name for CP chart
cp_chc_chart_html_file = 'REML1C_CP-Plot.html'   # HTML filename for CP chart
cp_chc_chart_trace_count = 2    # no. of traces in CP chart
er_cha_pr_chart_title = 'PR ER Plot for REML1A'  # title for ER chart
er_cha_pr_chart_xlabel = 'Date'        # xaxis name for ER chart
er_cha_pr_chart_ylabel = 'PR ER (A/min)'   # yaxis name for ER chart
er_cha_pr_chart_html_file = 'REML1A_PR_ER-Plot.html'   # HTML filename for ER chart
er_cha_pr_chart_trace_count = 5    # no. of traces in Nit ER chart
unif_cha_pr_chart_title = 'PR Uniformity Plot for REML1A'  # title for Unif chart
unif_cha_pr_chart_xlabel = 'Date'      # xaxis name for Unif chart
unif_cha_pr_chart_ylabel = 'PR Unif (%)'    # yaxis name for Unif chart
unif_cha_pr_chart_html_file = 'REML1A_PR_Unif-Plot.html'   # HTML filename for Unif chart
unif_cha_pr_chart_trace_count = 3    # no. of traces in Nit Unif chart
er_chc_pr_chart_title = 'PR ER Plot for REML1C'  # title for ER chart
er_chc_pr_chart_xlabel = 'Date'        # xaxis name for ER chart
er_chc_pr_chart_ylabel = 'PR ER (A/min)'   # yaxis name for ER chart
er_chc_pr_chart_html_file = 'REML1C_PR_ER-Plot.html'   # HTML filename for ER chart
er_chc_pr_chart_trace_count = 5    # no. of traces in ER chart
unif_chc_pr_chart_title = 'PR Uniformity Plot for REML1C'  # title for Unif chart
unif_chc_pr_chart_xlabel = 'Date'      # xaxis name for Unif chart
unif_chc_pr_chart_ylabel = 'PR Unif (%)'    # yaxis name for Unif chart
unif_chc_pr_chart_html_file = 'REML1C_PR_Unif-Plot.html'   # HTML filename for Unif chart
unif_chc_pr_chart_trace_count = 3    # no. of traces in Unif chart

excel_file_directory = "I:\\github_repos\\AutoPlot\\macro_enabled_logbooks\\CNT02_QC_LOG_BOOK\\CNT02_QC_LOG_BOOK.xlsm"
# excel_file_directory = "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\CNT_02_LOG_BOOK\\CNT02_QC_LOG_BOOK_macro\\CNT02_QC_LOG_BOOK.xlsm"

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots CP Chart with `cp_chart_trace_count` traces v/s Date.
"x": Date (x-axis) for CP Chart
"y1": Delta-CP (y-axis) for CP Chart
"y2": USL (y-axis) for CP Chart
# "y3": UCL (y-axis) for CP Chart
"""
def draw_plotly_reml1a_cp_plot(x, y1, y2, remarks):
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
            title = cp_cha_chart_title,
            xaxis = dict(title= cp_cha_chart_xlabel),
            yaxis = dict(title= cp_cha_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= cp_cha_chart_html_file)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------

"""
"Description": This function plots CP Chart with `cp_chart_trace_count` traces v/s Date.
"x": Date (x-axis) for CP Chart
"y1": Delta-CP (y-axis) for CP Chart
"y2": USL (y-axis) for CP Chart
# "y3": UCL (y-axis) for CP Chart
"""
def draw_plotly_reml1c_cp_plot(x, y1, y2, remarks):
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
            title = cp_chc_chart_title,
            xaxis = dict(title= cp_chc_chart_xlabel),
            yaxis = dict(title= cp_chc_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= cp_chc_chart_html_file)

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
def draw_plotly_reml1a_er_pr_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_cha_pr_chart_title,
            xaxis = dict(title= er_cha_pr_chart_xlabel),
            yaxis = dict(title= er_cha_pr_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= er_cha_pr_chart_html_file)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with `unif_nit_chart_trace_count` traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_reml1a_unif_pr_plot(x, y1, y2, y3, remarks):
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
            title = unif_cha_pr_chart_title,
            xaxis = dict(title= unif_cha_pr_chart_xlabel),
            yaxis = dict(title= unif_cha_pr_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= unif_cha_pr_chart_html_file)

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
def draw_plotly_reml1c_er_pr_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_chc_pr_chart_title,
            xaxis = dict(title= er_chc_pr_chart_xlabel),
            yaxis = dict(title= er_chc_pr_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= er_chc_pr_chart_html_file)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with `unif_nit_chart_trace_count` traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
# "y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_reml1c_unif_pr_plot(x, y1, y2, remarks):
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
            title = unif_chc_pr_chart_title,
            xaxis = dict(title= unif_chc_pr_chart_xlabel),
            yaxis = dict(title= unif_chc_pr_chart_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= unif_chc_pr_chart_html_file)

#====================================================================================================================================================================
#####################################################################################################################################################################
def main():
    wb = xw.Book.caller()
    # wb.sheets[0].range("A1").value = "Hello xlwings!"		# test code

    #****************************************************************************************************************************************************************
    # Define sheets
    sht_reml1_cp = wb.sheets['REML1-CP']
    sht_reml1a_er_pr = wb.sheets['PR Ch A ER']
    sht_reml1c_er_pr = wb.sheets['PR Ch C ER']
    # sht_reml1_plot_cp = wb.sheets['CP Plot']
    # sht_reml1_plot_er_ch_a_pr = wb.sheets['PR Ch A Plot']
    # sht_reml1_plot_er_ch_c_pr = wb.sheets['PR Ch C Plot']

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for CP
    df_reml1_cp = sht_reml1_cp.range('A9').options(
        pd.DataFrame, header=1, index=False, expand='table'
        ).value											                # fetch the data from sheet- 'ASBE1-CP'
    df_reml1_cp['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_reml1_cp = df_reml1_cp[["Date (MM/DD/YYYY)", "Chamber", "delta CP", "USL", "Remarks"]]        # The final dataframe with required columns
    df_reml1_cp = df_reml1_cp.dropna()                                              # dropping rows where at least one element is missing
    # sht_reml1_plot_cp.range('A46').options(index=False).value = df_reml1_cp   	    # show the dataframe values into sheet- 'CP Plot'
    df_reml1a_cp = df_reml1_cp[df_reml1_cp["Chamber"] == 'DPS']					# dataframe for DPS (chamber A)
    # sht_reml1_plot_cp.range('A46').options(index=False).value = df_reml1a_cp           # show the dataframe values into sheet- 'CP Plot'
    df_reml1c_cp = df_reml1_cp[df_reml1_cp["Chamber"] == 'ASP']					# dataframe for ASP (chamber C)
    # sht_reml1_plot_cp.range('G46').options(index=False).value = df_reml1c_cp           # show the dataframe values into sheet- 'CP Plot'
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_reml1a_cp_date = df_reml1a_cp["Date (MM/DD/YYYY)"]
    df_reml1a_cp_delta_cp = df_reml1a_cp["delta CP"]
    df_reml1a_cp_usl = df_reml1a_cp["USL"]
    # df_reml1a_cp_ucl = df_reml1a_cp["UCL"]
    df_reml1a_cp_remarks = df_reml1a_cp["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_reml1c_cp_date = df_reml1c_cp["Date (MM/DD/YYYY)"]
    df_reml1c_cp_delta_cp = df_reml1c_cp["delta CP"]
    df_reml1c_cp_usl = df_reml1c_cp["USL"]
    # df_reml1c_cp_ucl = df_reml1c_cp["UCL"]
    df_reml1c_cp_remarks = df_reml1c_cp["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # # Draw CP Plot for DPS chamber i.e. Ch-A
    # fig_reml1_cp_ch_a, ax_reml1_cp_ch_a = plt.subplots(1,1, figsize=(20,6))
    # monthyearFmt_cp_ch_a = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    # ax_reml1_cp_ch_a.xaxis.set_major_formatter(monthyearFmt_cp_ch_a)
    # _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # # ax_reml1_cp_ch_a.xaxis.set_major_locator(mdates.MonthLocator())          # set ticks after every Month
    # ax_reml1_cp_ch_a.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    # ax_reml1_cp_ch_a.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    # plt.plot(df_reml1a_cp["Date (MM/DD/YYYY)"], df_reml1a_cp["delta CP"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs CP
    # plt.plot(df_reml1a_cp["Date (MM/DD/YYYY)"], df_reml1a_cp["USL"], linestyle='-', color='#0000CD')        # plot date vs USL
    # plt.xlabel('Date', fontsize=18)      # xlabel
    # plt.ylabel('DPS delta CP (no.s)', fontsize=18)     # ylabel
    # # Custom Legends
    # custom_lines_cp_ch_a = [
    #     Line2D([0], [0], color='#FF7F50', lw=4),
    #     Line2D([0], [0], color='#0000CD', lw=4),
    #     ]
    # ax_reml1_cp_ch_a.legend(custom_lines_cp_ch_a, ['CP', 'USL'], fontsize=11, loc='upper right')  
    # lines_cp_ch_a = ax_reml1_cp_ch_a.plot(df_reml1a_cp["Date (MM/DD/YYYY)"], df_reml1a_cp["delta CP"], visible=False)
    # # datacursor(lines_cp_ch_a, hover=True, point_labels=df_reml1a_cp['Remarks'])
    # # plt.show()
    # # sht_cp_plot.activate()
    # # pic_cp = plt.show()
    # # plt.show('ASBE1_CP_Plot', left=xw.Range('A1').left, top=xw.Range('A1').top)      # this would activate hover 
    # # sh_plot_cp.pictures.add(pic_cp, name= "ASFE1_CP_Plot", update= True)
    # sht_reml1_plot_cp.pictures.add(fig_reml1_cp_ch_a, name= "REML1_DPS_CP_Plot", update= True)


    # Draw CP Plot (using Plotly) in Browser 
    draw_plotly_reml1a_cp_plot(
        x = df_reml1a_cp_date, 
        y1 = df_reml1a_cp_delta_cp, 
        y2 = df_reml1a_cp_usl, 
        # y3 = df_reml1a_cp_ucl,
        remarks = df_reml1a_cp_remarks
        )

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # # Draw CP Plot for ASP chamber i.e. Ch-C
    # fig_reml1_cp_ch_c, ax_reml1_cp_ch_c = plt.subplots(1,1, figsize=(20,6))
    # monthyearFmt_cp_ch_c = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    # ax_reml1_cp_ch_c.xaxis.set_major_formatter(monthyearFmt_cp_ch_c)
    # _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # # ax_reml1_cp_ch_c.xaxis.set_major_locator(mdates.MonthLocator())          # set ticks after every Month
    # ax_reml1_cp_ch_c.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    # ax_reml1_cp_ch_c.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    # plt.plot(df_reml1c_cp["Date (MM/DD/YYYY)"], df_reml1c_cp["delta CP"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs CP
    # plt.plot(df_reml1c_cp["Date (MM/DD/YYYY)"], df_reml1c_cp["USL"], linestyle='-', color='#0000CD')        # plot date vs USL
    # plt.xlabel('Date', fontsize=18)      # xlabel
    # plt.ylabel('ASP delta CP (no.s)', fontsize=18)     # ylabel
    # # Custom Legends
    # custom_lines_cp_ch_c = [
    #     Line2D([0], [0], color='#FF7F50', lw=4),
    #     Line2D([0], [0], color='#0000CD', lw=4),
    #     ]
    # ax_reml1_cp_ch_c.legend(custom_lines_cp_ch_c, ['CP', 'USL'], fontsize=11, loc='upper right')  
    # lines_cp_ch_c = ax_reml1_cp_ch_c.plot(df_reml1c_cp["Date (MM/DD/YYYY)"], df_reml1c_cp["delta CP"], visible=False)
    # # datacursor(lines_cp_ch_c, hover=True, point_labels=df_reml1c_cp['Remarks'])
    # # plt.show()
    # # sht_cp_plot.activate()
    # # pic_cp = plt.show()
    # # plt.show('ASBE1_CP_Plot', left=xw.Range('A1').left, top=xw.Range('A1').top)      # this would activate hover 
    # # sht_reml1_plot_cp.pictures.add(pic_cp, name= "ASFE1_CP_Plot", update= True)
    # sht_reml1_plot_cp.pictures.add(fig_reml1_cp_ch_c, name= "REML1_ASP_CP_Plot", update= True)

    # Draw CP Plot (using Plotly) in Browser 
    draw_plotly_reml1c_cp_plot(
        x = df_reml1c_cp_date, 
        y1 = df_reml1c_cp_delta_cp, 
        y2 = df_reml1c_cp_usl, 
        # y3 = df_reml1c_cp_ucl,
        remarks = df_reml1c_cp_remarks
        )

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for Ch A PR i.e. DPS chamber ER & Unif PLot
    # data_folder = Path(os.getcwd())
    # file_to_open = data_folder / "ASH09_QC_LOG_BOOK.xlsm"
    # excel_file = pd.ExcelFile(file_to_open)

    excel_file_sht_reml1a_pr = pd.ExcelFile(excel_file_directory)
    df_reml1a_er_pr = excel_file_sht_reml1a_pr.parse('PR Ch A ER', skiprows=8)                            # copy a sheet and paste into another sheet and skiprows 8
    
    df_reml1a_er_pr = df_reml1a_er_pr[["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "LSL", "USL", "LCL", "UCL", "Remarks", "% Uni USL", "% Uni UCL"]]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_reml1a_er_pr['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_reml1a_er_pr = df_reml1a_er_pr.dropna()                                              # dropping rows where at least one element is missing
    # sht_plot_nit.range('A28').options(index=False).value = df_reml1a_er_pr        # show the dataframe values into sheet- 'CP Plot'
    
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for REML1A PR ER & Unif PLot
    df_reml1a_er_pr_date = df_reml1a_er_pr["Date (MM/DD/YYYY)"]
    df_reml1a_er_pr_er = df_reml1a_er_pr["Etch Rate (A/Min)"]
    df_reml1a_er_pr_usl = df_reml1a_er_pr["USL"]
    df_reml1a_er_pr_lsl = df_reml1a_er_pr["LSL"]
    df_reml1a_er_pr_ucl = df_reml1a_er_pr["UCL"]
    df_reml1a_er_pr_lcl = df_reml1a_er_pr["LCL"]
    df_reml1a_er_pr_unif = df_reml1a_er_pr["% Uni"]
    df_reml1a_er_pr_unif_usl = df_reml1a_er_pr["% Uni USL"]
    df_reml1a_er_pr_unif_ucl = df_reml1a_er_pr["% Uni UCL"]
    df_reml1a_er_pr_remarks = df_reml1a_er_pr["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # # Draw Ch A PR ER PLot
    # fig_reml1_er_ch_a_pr, ax_reml1_er_ch_a_pr = plt.subplots(1,1, figsize=(20,6))
    # monthyearFmt_er_pr = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    # ax_reml1_er_ch_a_pr.xaxis.set_major_formatter(monthyearFmt_er_pr)
    # _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # # ax_reml1_er_ch_a_pr.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    # ax_reml1_er_ch_a_pr.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    # ax_reml1_er_ch_a_pr.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    # plt.plot(df_reml1a_er_pr["Date (MM/DD/YYYY)"], df_reml1a_er_pr["Etch Rate (A/Min)"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    # plt.plot(df_reml1a_er_pr["Date (MM/DD/YYYY)"], df_reml1a_er_pr["USL"], linestyle='-', color='#0000CD')        # plot date vs USL
    # plt.plot(df_reml1a_er_pr["Date (MM/DD/YYYY)"], df_reml1a_er_pr["LSL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # plt.plot(df_reml1a_er_pr["Date (MM/DD/YYYY)"], df_reml1a_er_pr["UCL"], linestyle='-', color='#FF1493')        # plot date vs UCL
    # plt.plot(df_reml1a_er_pr["Date (MM/DD/YYYY)"], df_reml1a_er_pr["LCL"], linestyle='-', color='#FF1493')        # plot date vs LCL
    # plt.xlabel('Date', fontsize=18)      # xlabel
    # plt.ylabel('Ch A PR ER (A/min)', fontsize=18)     # ylabel
    # # Custom Legends
    # custom_lines_er_cha_pr = [
    #     Line2D([0], [0], color='#FF7F50', lw=4),	# ER
    #     Line2D([0], [0], color='#0000CD', lw=4),	# USL
    #     Line2D([0], [0], color='#0000CD', lw=4),	# LSL
    #     Line2D([0], [0], color='#FF1493', lw=4),	# UCL
    #     Line2D([0], [0], color='#FF1493', lw=4)		# LCL        
    #     ]
    # ax_reml1_er_ch_a_pr.legend(custom_lines_er_cha_pr, ['ER', 'USL', 'LSL', 'UCL', 'LCL'], fontsize=11, loc='upper right') 
    # lines_er_ch_a_pr = ax_reml1_er_ch_a_pr.plot(df_reml1a_er_pr["Date (MM/DD/YYYY)"], df_reml1a_er_pr["Etch Rate (A/Min)"], visible=False)
    # # datacursor(lines_er_ch_a_pr, hover=True, point_labels=df_reml1a_er_pr['Remarks'])
    # # plt.show()        # shows 2 figures in different windows
    # sht_reml1_plot_er_ch_a_pr.pictures.add(fig_reml1_er_ch_a_pr, name= "REML1_CH_A_PR_ER_Plot", update= True)

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
	# # Draw ChA PR Unif PLot
 #    fig_reml1_unif_ch_a_pr, ax_reml1_unif_ch_a_pr = plt.subplots(1,1, figsize=(20,6))
 #    monthyearFmt_unif_ch_a_pr = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
 #    ax_reml1_unif_ch_a_pr.xaxis.set_major_formatter(monthyearFmt_unif_ch_a_pr)
 #    _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
 #    # ax_reml1_unif_ch_a_pr.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
 #    ax_reml1_unif_ch_a_pr.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
 #    ax_reml1_unif_ch_a_pr.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
 #    plt.plot(df_reml1a_er_pr["Date (MM/DD/YYYY)"], df_reml1a_er_pr["% Uniformity"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
 #    plt.plot(df_reml1a_er_pr["Date (MM/DD/YYYY)"], df_reml1a_er_pr["% Uni USL"], linestyle='-', color='#0000CD')        # plot date vs USL
 #    plt.plot(df_reml1a_er_pr["Date (MM/DD/YYYY)"], df_reml1a_er_pr["% Uni UCL"], linestyle='-', color='#FF1493')        # plot date vs UCL
 #    plt.xlabel('Date', fontsize=18)      # xlabel
 #    plt.ylabel('Ch A PR Unif (%)', fontsize=18)     # ylabel
 #    # Custom Legends
 #    custom_lines_unif_ch_a_pr = [
 #        Line2D([0], [0], color='#FF7F50', lw=4),	# ER
 #        Line2D([0], [0], color='#0000CD', lw=4),	# USL
 #        Line2D([0], [0], color='#FF1493', lw=4),	# UCL
 #        ]
 #    ax_reml1_unif_ch_a_pr.legend(custom_lines_unif_ch_a_pr, ['Unif', 'USL', 'UCL'], fontsize=11, loc='upper right') 
 #    lines_unif_ch_a_pr = ax_reml1_unif_ch_a_pr.plot(df_reml1a_er_pr["Date (MM/DD/YYYY)"], df_reml1a_er_pr["% Uniformity"], visible=False)
 #    # datacursor(lines_unif_ch_a_pr, hover=True, point_labels=df_reml1a_er_pr['Remarks'])
 #    # plt.show()        # shows 2 figures in different windows
 #    sht_reml1_plot_er_ch_a_pr.pictures.add(fig_reml1_unif_ch_a_pr, name= "REML1_CH_A_PR_UNIF_Plot", update= True)

    # Draw REML1A PR ER Plot (using Plotly) in Browser 
    draw_plotly_reml1a_er_pr_plot(
        x = df_reml1a_er_pr_date, 
        y1 = df_reml1a_er_pr_er,
        y2 = df_reml1a_er_pr_usl, 
        y3 = df_reml1a_er_pr_lsl,
        y4 = df_reml1a_er_pr_ucl,
        y5 = df_reml1a_er_pr_lcl,
        remarks = df_reml1a_er_pr_remarks
        )
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw REML1A PR Unif Plot (using Plotly) in Browser     
    draw_plotly_reml1a_unif_pr_plot(
        x = df_reml1a_er_pr_date, 
        y1 = df_reml1a_er_pr_unif, 
        y2 = df_reml1a_er_pr_unif_usl,
        y3 = df_reml1a_er_pr_unif_ucl,
        remarks = df_reml1a_er_pr_remarks
        )


    #****************************************************************************************************************************************************************
    # Fetch Dataframe for Ch C PR i.e. ASP Chamber ER & Unif PLot
    # data_folder = Path(os.getcwd())
    # file_to_open = data_folder / "ASH09_QC_LOG_BOOK.xlsm"
    # excel_file = pd.ExcelFile(file_to_open)

    excel_file_sht_reml1c_pr = pd.ExcelFile(excel_file_directory)
    df_reml1c_er_pr = excel_file_sht_reml1c_pr.parse('PR Ch C ER', skiprows=8)                            # copy a sheet and paste into another sheet and skiprows 8
    
    df_reml1c_er_pr = df_reml1c_er_pr[["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "LSL", "USL", "LCL", "UCL", "Remarks", "% Uni USL"]]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_reml1c_er_pr['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_reml1c_er_pr = df_reml1c_er_pr.dropna()                                              # dropping rows where at least one element is missing
    # sht_reml1_plot_er_ch_c_pr.range('A28').options(index=False).value = df_reml1c_er_pr        # show the dataframe values into sheet- 'CP Plot'
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for REML1C PR ER & Unif PLot
    df_reml1c_er_pr_date = df_reml1c_er_pr["Date (MM/DD/YYYY)"]
    df_reml1c_er_pr_er = df_reml1c_er_pr["Etch Rate (A/Min)"]
    df_reml1c_er_pr_usl = df_reml1c_er_pr["USL"]
    df_reml1c_er_pr_lsl = df_reml1c_er_pr["LSL"]
    df_reml1c_er_pr_ucl = df_reml1c_er_pr["UCL"]
    df_reml1c_er_pr_lcl = df_reml1c_er_pr["LCL"]
    df_reml1c_er_pr_unif = df_reml1c_er_pr["% Uni"]
    df_reml1c_er_pr_unif_usl = df_reml1c_er_pr["% Uni USL"]
    # df_reml1c_er_pr_unif_ucl = df_reml1c_er_pr["% Uni UCL"]
    df_reml1c_er_pr_remarks = df_reml1c_er_pr["Remarks"]
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # # Draw ChC ER PLot
    # fig_reml1_er_ch_c_pr, ax_reml1_er_ch_c_pr = plt.subplots(1,1, figsize=(20,6))
    # monthyearFmt_er_ch_c_pr = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    # ax_reml1_er_ch_c_pr.xaxis.set_major_formatter(monthyearFmt_er_ch_c_pr)
    # _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # # ax_reml1_er_ch_c_pr.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    # ax_reml1_er_ch_c_pr.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    # ax_reml1_er_ch_c_pr.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    # plt.plot(df_reml1c_er_pr["Date (MM/DD/YYYY)"], df_reml1c_er_pr["Etch Rate (A/Min)"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    # plt.plot(df_reml1c_er_pr["Date (MM/DD/YYYY)"], df_reml1c_er_pr["USL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # plt.plot(df_reml1c_er_pr["Date (MM/DD/YYYY)"], df_reml1c_er_pr["LSL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # plt.plot(df_reml1c_er_pr["Date (MM/DD/YYYY)"], df_reml1c_er_pr["UCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    # plt.plot(df_reml1c_er_pr["Date (MM/DD/YYYY)"], df_reml1c_er_pr["LCL"], linestyle='-', color='#FF1493')        # plot date vs LSL
    # plt.xlabel('Date', fontsize=18)      # xlabel
    # plt.ylabel('Ch C PR ER (A/min)', fontsize=18)     # ylabel
    # # Custom Legends
    # custom_lines_er_ch_c_pr = [
    #     Line2D([0], [0], color='#FF7F50', lw=4),	# ER
    #     Line2D([0], [0], color='#0000CD', lw=4),	# USL
    #     Line2D([0], [0], color='#0000CD', lw=4),	# LSL
    #     Line2D([0], [0], color='#FF1493', lw=4),	# UCL
    #     Line2D([0], [0], color='#FF1493', lw=4)		# LCL        
    #     ]
    # ax_reml1_er_ch_c_pr.legend(custom_lines_er_ch_c_pr, ['ER', 'USL', 'LSL', 'UCL', 'LCL'], fontsize=11, loc='upper right') 
    # lines_er_ch_c_pr = ax_reml1_er_ch_c_pr.plot(df_reml1c_er_pr["Date (MM/DD/YYYY)"], df_reml1c_er_pr["Etch Rate (A/Min)"], visible=False)
    # # datacursor(lines_er_ch_c_pr, hover=True, point_labels=df_reml1c_er_pr['Remarks'])
    # # plt.show()        # shows 2 figures in different windows
    # sht_reml1_plot_er_ch_c_pr.pictures.add(fig_reml1_er_ch_c_pr, name= "REML1_CH_C_PR_ER_Plot", update= True)

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # # Draw ChC UNIF PLot
    # fig_reml1_unif_ch_c_pr, ax_reml1_unif_ch_c_pr = plt.subplots(1,1, figsize=(20,6))
    # monthyearFmt_unif_ch_c_pr = mdates.DateFormatter('%Y-%b-%d')                        # formatting as 2017-Jan-14
    # ax_reml1_unif_ch_c_pr.xaxis.set_major_formatter(monthyearFmt_unif_ch_c_pr)
    # _ = plt.xticks(rotation=90)                                         # rotating 90 counterclockwise
    # # ax_reml1_unif_ch_c_pr.xaxis.set_major_locator(mdates.MonthLocator())         # set ticks after every 2 Mondays
    # ax_reml1_unif_ch_c_pr.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=MO, interval=2))          # set ticks after every 2 Mondays
    # ax_reml1_unif_ch_c_pr.grid(which='both', alpha=0.15)           # set grid with transparency to 0.15
    # plt.plot(df_reml1c_er_pr["Date (MM/DD/YYYY)"], df_reml1c_er_pr["% Uniformity"], linestyle='-', marker='o', markerfacecolor='#008000', color='#FF7F50')    # plot date vs ER
    # plt.plot(df_reml1c_er_pr["Date (MM/DD/YYYY)"], df_reml1c_er_pr["% Uni USL"], linestyle='-', color='#0000CD')        # plot date vs LSL
    # plt.xlabel('Date', fontsize=18)      # xlabel
    # plt.ylabel('Ch C PR Unif (%)', fontsize=18)     # ylabel
    # # Custom Legends
    # custom_lines_unif_ch_c_pr = [
    #     Line2D([0], [0], color='#FF7F50', lw=4),	# ER
    #     Line2D([0], [0], color='#0000CD', lw=4),	# USL
    #     Line2D([0], [0], color='#FF1493', lw=4),	# UCL
    #     ]
    # ax_reml1_unif_ch_c_pr.legend(custom_lines_unif_ch_c_pr, ['Unif', 'USL'], fontsize=11, loc='upper right') 
    # lines_unif_ch_c_pr = ax_reml1_unif_ch_c_pr.plot(df_reml1c_er_pr["Date (MM/DD/YYYY)"], df_reml1c_er_pr["% Uniformity"], visible=False)
    # # datacursor(lines_unif_ch_c_pr, hover=True, point_labels=df_reml1c_er_pr['Remarks'])
    # # plt.show()        # shows 2 figures in different windows
    # sht_reml1_plot_er_ch_c_pr.pictures.add(fig_reml1_unif_ch_c_pr, name= "REML1_CH_C_PR_UNIF_Plot", update= True)

    # Draw REML1C PR ER Plot (using Plotly) in Browser 
    draw_plotly_reml1c_er_pr_plot(
        x = df_reml1c_er_pr_date, 
        y1 = df_reml1c_er_pr_er,
        y2 = df_reml1c_er_pr_usl, 
        y3 = df_reml1c_er_pr_lsl,
        y4 = df_reml1c_er_pr_ucl,
        y5 = df_reml1c_er_pr_lcl,
        remarks = df_reml1c_er_pr_remarks
        )
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw REML1C PR Unif Plot (using Plotly) in Browser     
    draw_plotly_reml1c_unif_pr_plot(
        x = df_reml1c_er_pr_date, 
        y1 = df_reml1c_er_pr_unif, 
        y2 = df_reml1c_er_pr_unif_usl,
        # y3 = df_reml1c_er_pr_unif_ucl,
        remarks = df_reml1c_er_pr_remarks
        )

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
# User Defined Functions (UDFs)
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
@xw.func
def hello(name):
    return "hello {0}".format(name)



