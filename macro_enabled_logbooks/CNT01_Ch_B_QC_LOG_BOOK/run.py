# Import packages
# import xlwings as xw
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
cp_plot_title = 'CP Plot for REPL1B'  # title for CP plot
cp_plot_xlabel = 'Date'   # xaxis name for CP plot
cp_plot_ylabel = 'delta CP (no.s)'     # yaxis name for CP plot
cp_plot_html_file = 'REPL1B_CP-Plot.html'   # HTML filename for CP plot
cp_plot_trace_count = 2    # no. of traces in CP plot
er_nit_plot_title = 'Nit ER Plot for REPL1B'  # title for Nit ER plot
er_nit_plot_xlabel = 'Date'        # xaxis name for Nit ER plot
er_nit_plot_ylabel = 'Nit ER (A/min)'   # yaxis name for Nit ER plot
er_nit_plot_html_file = 'REPL1B_Nit_ER-Plot.html'   # HTML filename for Nit ER plot
er_nit_plot_trace_count = 5    # no. of traces in Nit ER plot
unif_nit_plot_title = 'Nit Uniformity Plot for REPL1B'  # title for Nit Unif plot
unif_nit_plot_xlabel = 'Date'      # xaxis name for Nit Unif plot
unif_nit_plot_ylabel = 'Nit Unif (%)'    # yaxis name for Nit Unif plot
unif_nit_plot_html_file = 'REPL1B_Nit_Unif-Plot.html'   # HTML filename for Nit Unif plot
unif_nit_plot_trace_count = 3    # no. of traces in Nit Unif plot
er_poly_plot_title = 'Poly ER Plot for REPL1B'  # title for Poly ER plot
er_poly_plot_xlabel = 'Date'        # xaxis name for Poly ER plot
er_poly_plot_ylabel = 'Poly ER (A/min)'   # yaxis name for Poly ER plot
er_poly_plot_html_file = 'REPL1B_Poly_ER-Plot.html'   # HTML filename for Poly ER plot
er_poly_plot_trace_count = 5    # no. of traces in Poly ER plot
unif_poly_plot_title = 'Poly Uniformity Plot for REPL1B'  # title for Poly Unif plot
unif_poly_plot_xlabel = 'Date'      # xaxis name for Poly Unif plot
unif_poly_plot_ylabel = 'Poly Unif (%)'    # yaxis name for Poly Unif plot
unif_poly_plot_html_file = 'REPL1B_Poly_Unif-Plot.html'   # HTML filename for Poly Unif plot
unif_poly_plot_trace_count = 3    # no. of traces in Poly Unif plot

excel_file_directory = "I:\\github_repos\\AutoPlot\\macro_enabled_logbooks\\CNT01_Ch_B_QC_LOG_BOOK\\CNT01_Ch_B_QC_LOG_BOOK.xlsm"
# excel_file_directory = "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\CNT_01_LOG_BOOK\\CNT01_QC_LOG_BOOK_Ch_A_macro\\CNT01_Ch_B_QC_LOG_BOOK.xlsm"

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
"Description": This function plots CP Chart with `cp_plot_trace_count` traces v/s Date.
"x": Date (x-axis) for CP Chart
"y1": Delta-CP (y-axis) for CP Chart
"y2": USL (y-axis) for CP Chart
# "y3": UCL (y-axis) for CP Chart
"""
def draw_plotly_repl1b_cp_plot(x, y1, y2, remarks):
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
            title = cp_plot_title,
            xaxis = dict(title= cp_plot_xlabel),
            yaxis = dict(title= cp_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= cp_plot_html_file, auto_open = False)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots ER Chart with `er_nit_plot_trace_count` traces v/s Date.
"x": Date (x-axis) for ER Chart
"y1": ER (y-axis) for ER Chart
"y2": USL (y-axis) for ER Chart
"y3": LSL (y-axis) for ER Chart
"y4": UCL (y-axis) for ER Chart
"y5": LCL (y-axis) for ER Chart
"""
def draw_plotly_repl1b_er_nit_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_nit_plot_title,
            xaxis = dict(title= er_nit_plot_xlabel),
            yaxis = dict(title= er_nit_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= er_nit_plot_html_file, auto_open = False)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with `unif_nit_plot_trace_count` traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_repl1b_unif_nit_plot(x, y1, y2, y3, remarks):
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
            title = unif_nit_plot_title,
            xaxis = dict(title= unif_nit_plot_xlabel),
            yaxis = dict(title= unif_nit_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= unif_nit_plot_html_file, auto_open = False)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots ER Chart with `er_poly_plot_trace_count` traces v/s Date.
"x": Date (x-axis) for ER Chart
"y1": ER (y-axis) for ER Chart
"y2": USL (y-axis) for ER Chart
"y3": LSL (y-axis) for ER Chart
"y4": UCL (y-axis) for ER Chart
"y5": LCL (y-axis) for ER Chart
"""
def draw_plotly_repl1b_er_poly_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_poly_plot_title,
            xaxis = dict(title= er_poly_plot_xlabel),
            yaxis = dict(title= er_poly_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= er_poly_plot_html_file, auto_open = False)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with `unif_poly_plot_trace_count` traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_repl1b_unif_poly_plot(x, y1, y2, y3, remarks):
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
            title = unif_poly_plot_title,
            xaxis = dict(title= unif_poly_plot_xlabel),
            yaxis = dict(title= unif_poly_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= unif_poly_plot_html_file, auto_open = False)


#====================================================================================================================================================================
#####################################################################################################################################################################
def main():
    # wb = xw.Book.caller()
    # wb.sheets[0].range("A1").value = "Hello xlwings!"     # test code

    #****************************************************************************************************************************************************************
    # Define sheets
    # sht_repl1b_cp = wb.sheets['REPL1B-CP']
    # sht_repl1b_er_nit = wb.sheets['REPL1B-ERNit']
    # sht_repl1b_er_poly = wb.sheets['REPL1B-ERPoly']
    # sht_repl1b_plot_cp = wb.sheets['CP Plot']
    # sht_repl1b_plot_er_nit = wb.sheets['Nit Plot']
    # sht_repl1b_plot_er_poly = wb.sheets['Poly Plot']
    excel_file = pd.ExcelFile(excel_file_directory)

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for CP Plot
    df_repl1b_cp = excel_file.parse('REPL1B-CP', skiprows=8)                            # copy a sheet and paste into another sheet and skiprows 9    
    df_repl1b_cp['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_repl1b_cp = df_repl1b_cp[["Date (MM/DD/YYYY)", "delta CP", "USL", "Remarks"]]        # The final dataframe with required columns
    df_repl1b_cp = df_repl1b_cp.dropna()                                              # dropping rows where at least one element is missing
    # sht_repl1b_plot_cp.range('A25').options(index=False).value = df_repl1b_cp         # show the dataframe values into sheet- 'CP Plot'
    
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_repl1b_cp_date = df_repl1b_cp["Date (MM/DD/YYYY)"]
    df_repl1b_cp_delta_cp = df_repl1b_cp["delta CP"]
    df_repl1b_cp_usl = df_repl1b_cp["USL"]
    # df_repl1b_cp_ucl = df_repl1b_cp["UCL"]
    df_repl1b_cp_remarks = df_repl1b_cp["Remarks"]


    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw CP Plot (using Plotly) in Browser 
    draw_plotly_repl1b_cp_plot(
        x = date_formatter(df_repl1b_cp_date), 
        y1 = df_repl1b_cp_delta_cp, 
        y2 = df_repl1b_cp_usl, 
        # y3 = df_repl1b_cp_ucl,
        remarks = df_repl1b_cp_remarks
        )

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for NIT ER & Unif Plot  
    # data_folder = Path(os.getcwd())
    # file_to_open = data_folder / "ASH09_QC_LOG_BOOK.xlsm"
    # excel_file = pd.ExcelFile(file_to_open)

    df_repl1b_er_nit = excel_file.parse('REPL1B-ERNit', skiprows=9)                            # copy a sheet and paste into another sheet and skiprows 9    
    df_repl1b_er_nit = df_repl1b_er_nit[["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "LSL", "USL", "LCL", "UCL", "Remarks", "% Uni USL", "% Uni UCL"]]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_repl1b_er_nit['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_repl1b_er_nit = df_repl1b_er_nit.dropna()                                              # dropping rows where at least one element is missing
    # sht_repl1b_plot_er_nit.range('A28').options(index=False).value = df_repl1b_er_nit        # show the dataframe values into sheet- 'CP Plot'
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param for Nit ER & Unif PLot
    df_repl1b_er_nit_date = df_repl1b_er_nit["Date (MM/DD/YYYY)"]
    df_repl1b_er_nit_er = df_repl1b_er_nit["Etch Rate (A/Min)"]
    df_repl1b_er_nit_usl = df_repl1b_er_nit["USL"]
    df_repl1b_er_nit_lsl = df_repl1b_er_nit["LSL"]
    df_repl1b_er_nit_ucl = df_repl1b_er_nit["UCL"]
    df_repl1b_er_nit_lcl = df_repl1b_er_nit["LCL"]
    df_repl1b_er_nit_unif = df_repl1b_er_nit["% Uni"]
    df_repl1b_er_nit_unif_usl = df_repl1b_er_nit["% Uni USL"]
    df_repl1b_er_nit_unif_ucl = df_repl1b_er_nit["% Uni UCL"]
    df_repl1b_er_nit_remarks = df_repl1b_er_nit["Remarks"]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw Nit ER Plot (using Plotly) in Browser 
    draw_plotly_repl1b_er_nit_plot(
        x = date_formatter(df_repl1b_er_nit_date), 
        y1 = df_repl1b_er_nit_er,
        y2 = df_repl1b_er_nit_usl, 
        y3 = df_repl1b_er_nit_lsl,
        y4 = df_repl1b_er_nit_ucl,
        y5 = df_repl1b_er_nit_lcl,
        remarks = df_repl1b_er_nit_remarks
        )
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw Nit Unif Plot (using Plotly) in Browser     
    draw_plotly_repl1b_unif_nit_plot(
        x = date_formatter(df_repl1b_er_nit_date), 
        y1 = df_repl1b_er_nit_unif, 
        y2 = df_repl1b_er_nit_unif_usl,
        y3 = df_repl1b_er_nit_unif_ucl,
        remarks = df_repl1b_er_nit_remarks
        )



    #****************************************************************************************************************************************************************
    # Fetch Dataframe for POLY ER & Unif PLot
    # data_folder = Path(os.getcwd())
    # file_to_open = data_folder / "ASH09_QC_LOG_BOOK.xlsm"
    # excel_file = pd.ExcelFile(file_to_open)

    df_repl1b_er_poly = excel_file.parse('REPL1B-ERPoly', skiprows=9)                            # copy a sheet and paste into another sheet and skiprows 9    
    df_repl1b_er_poly = df_repl1b_er_poly[["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "LSL", "USL", "LCL", "UCL", "Remarks", "% Uni USL", "% Uni UCL"]]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_repl1b_er_poly['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_repl1b_er_poly = df_repl1b_er_poly.dropna()                                              # dropping rows where at least one element is missing
    # sht_repl1b_plot_er_poly.range('A28').options(index=False).value = df_repl1b_er_poly        # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param for Nit ER & Unif PLot
    df_repl1b_er_poly_date = df_repl1b_er_poly["Date (MM/DD/YYYY)"]
    df_repl1b_er_poly_er = df_repl1b_er_poly["Etch Rate (A/Min)"]
    df_repl1b_er_poly_usl = df_repl1b_er_poly["USL"]
    df_repl1b_er_poly_lsl = df_repl1b_er_poly["LSL"]
    df_repl1b_er_poly_ucl = df_repl1b_er_poly["UCL"]
    df_repl1b_er_poly_lcl = df_repl1b_er_poly["LCL"]
    df_repl1b_er_poly_unif = df_repl1b_er_poly["% Uni"]
    df_repl1b_er_poly_unif_usl = df_repl1b_er_poly["% Uni USL"]
    df_repl1b_er_poly_unif_ucl = df_repl1b_er_poly["% Uni UCL"]
    df_repl1b_er_poly_remarks = df_repl1b_er_poly["Remarks"]
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw Poly ER Plot (using Plotly) in Browser 
    draw_plotly_repl1b_er_poly_plot(
        x = date_formatter(df_repl1b_er_poly_date), 
        y1 = df_repl1b_er_poly_er,
        y2 = df_repl1b_er_poly_usl, 
        y3 = df_repl1b_er_poly_lsl,
        y4 = df_repl1b_er_poly_ucl,
        y5 = df_repl1b_er_poly_lcl,
        remarks = df_repl1b_er_poly_remarks
        )
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw Poly Unif Plot (using Plotly) in Browser     
    draw_plotly_repl1b_unif_poly_plot(
        x = date_formatter(df_repl1b_er_poly_date), 
        y1 = df_repl1b_er_poly_unif, 
        y2 = df_repl1b_er_poly_unif_usl,
        y3 = df_repl1b_er_poly_unif_ucl,
        remarks = df_repl1b_er_poly_remarks
        )

if __name__ == '__main__':
    main()
