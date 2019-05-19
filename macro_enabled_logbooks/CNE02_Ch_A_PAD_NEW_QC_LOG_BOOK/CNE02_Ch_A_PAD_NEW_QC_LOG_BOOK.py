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
cp_plot_title = 'CP Plot for REOX1A'  # title for CP plot
cp_plot_xlabel = 'Date'   # xaxis name for CP plot
cp_plot_ylabel = 'delta CP (no.s)'     # yaxis name for CP plot
cp_plot_html_file = 'REOX1A_CP-Plot.html'   # HTML filename for CP plot
cp_plot_trace_count = 2    # no. of traces in CP plot
er_sin_plot_title = 'SiN ER Plot for REOX1A'  # title for ER plot
er_sin_plot_xlabel = 'Date'        # xaxis name for ER plot
er_sin_plot_ylabel = 'SiN ER (A/min)'   # yaxis name for ER plot
er_sin_plot_html_file = 'REOX1A_SiN_ER-Plot.html'   # HTML filename for ER plot
er_sin_plot_trace_count = 5    # no. of traces in ER plot
unif_sin_plot_title = 'SiN Uniformity Plot for REOX1A'  # title for Unif plot
unif_sin_plot_xlabel = 'Date'      # xaxis name for Unif plot
unif_sin_plot_ylabel = 'SiN Unif (%)'    # yaxis name for Unif plot
unif_sin_plot_html_file = 'REOX1A_SiN_Unif-Plot.html'   # HTML filename for Unif plot
unif_sin_plot_trace_count = 3    # no. of traces in Unif plot
er_teos_plot_title = 'TEOS ER Plot for REOX1A'  # title for ER plot
er_teos_plot_xlabel = 'Date'        # xaxis name for ER plot
er_teos_plot_ylabel = 'TEOS ER (A/min)'   # yaxis name for ER plot
er_teos_plot_html_file = 'REOX1A_TEOS_ER-Plot.html'   # HTML filename for ER plot
er_teos_plot_trace_count = 5    # no. of traces in ER plot
unif_teos_plot_title = 'TEOS Uniformity Plot for REOX1A'  # title for Unif plot
unif_teos_plot_xlabel = 'Date'      # xaxis name for Unif plot
unif_teos_plot_ylabel = 'TEOS Unif (%)'    # yaxis name for Unif plot
unif_teos_plot_html_file = 'REOX1A_TEOS_Unif-Plot.html'   # HTML filename for Unif plot
unif_teos_plot_trace_count = 3    # no. of traces in Unif plot
sht_cp_columns = ["Date (MM/DD/YYYY)", "delta CP", "USL", "Remarks"]
sht_er_columns = ["Date (MM/DD/YYYY)", "Layer", "Etch Rate (A/Min)", "% Uni", "Remarks", "LSL", "USL", "LCL", "UCL", "% Uni USL", "% Uni UCL"]

excel_file_directory = "I:\\github_repos\\AutoPlot\\macro_enabled_logbooks\\CNE02_Ch_A_PAD_NEW_QC_LOG_BOOK\\CNE02_Ch_A_PAD_NEW_QC_LOG_BOOK.xlsm"
# excel_file_directory = "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\UNT_02_LOG_BOOK\\CNE02_Ch_A_PAD_NEW_QC_LOG_BOOK\\CNE02_Ch_A_PAD_NEW_QC_LOG_BOOK.xlsm"

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
"Description": This function plots CP Chart with traces v/s Date.
"x": Date (x-axis) for CP Chart
"y1": Delta-CP (y-axis) for CP Chart
"y2": USL (y-axis) for CP Chart
# "y3": UCL (y-axis) for CP Chart
"""
def draw_plotly_reox1a_cp_plot(x, y1, y2, remarks):
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
    py.offline.plot(fig, filename= cp_plot_html_file)

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
def draw_plotly_reox1a_er_sin_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_sin_plot_title,
            xaxis = dict(title= er_sin_plot_xlabel),
            yaxis = dict(title= er_sin_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= er_sin_plot_html_file)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_reox1a_unif_sin_plot(x, y1, y2, y3, remarks):
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
            title = unif_sin_plot_title,
            xaxis = dict(title= unif_sin_plot_xlabel),
            yaxis = dict(title= unif_sin_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= unif_sin_plot_html_file)

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
def draw_plotly_reox1a_er_teos_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_teos_plot_title,
            xaxis = dict(title= er_teos_plot_xlabel),
            yaxis = dict(title= er_teos_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= er_teos_plot_html_file)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots Unif Chart with traces v/s Date.
"x": Date (x-axis) for Unif Chart
"y1": Unif (y-axis) for Unif Chart
"y2": USL (y-axis) for Unif Chart
"y3": UCL (y-axis) for Unif Chart
"""
def draw_plotly_reox1a_unif_teos_plot(x, y1, y2, y3, remarks):
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
            title = unif_teos_plot_title,
            xaxis = dict(title= unif_teos_plot_xlabel),
            yaxis = dict(title= unif_teos_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= unif_teos_plot_html_file)



#====================================================================================================================================================================
#####################################################################################################################################################################
def main():
    wb = xw.Book.caller()
    # wb.sheets[0].range("A1").value = "Hello xlwings!"     # test code

    #****************************************************************************************************************************************************************
    # Define sheets
    sht_reox1a_cp = wb.sheets['REOX1A-CP']
    sht_reox1a_er = wb.sheets['REOX1A-ER']
    # sht_reox1a_plot_cp = wb.sheets['CP Plot']
    # sht_reox1a_plot_barc = wb.sheets['BARC Plot']
    # sht_reox1a_plot_pr = wb.sheets['SiN Plot']
    # sht_reox1a_plot_teos = wb.sheets['TEOS Plot']
    # sht_reox1a_plot_sin = wb.sheets['SiN Plot']

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for CP
    df_reox1a_cp = sht_reox1a_cp.range('A9').options(
        pd.DataFrame, header=1, index=False, expand='table'
        ).value                                                         # fetch the data from sheet- 'REOX1A-CP'
    df_reox1a_cp = df_reox1a_cp[sht_cp_columns]        # The final dataframe with required columns
    df_reox1a_cp['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_reox1a_cp = df_reox1a_cp.dropna()                                              # dropping rows where at least one element is missing
    # sht_reox1a_plot_cp.range('A25').options(index=False).value = df_reox1a_cp         # show the dataframe values into sheet- 'CP Plot'
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
        x = date_formatter(df_reox1a_cp_date), 
        y1 = df_reox1a_cp_delta_cp, 
        y2 = df_reox1a_cp_usl, 
        # y3 = df_reox1a_cp_ucl,
        remarks = df_reox1a_cp_remarks
        )


    #****************************************************************************************************************************************************************
    """
    Fetch Dataframes for 
        - SiN, 
        - TEOS 
    """
    # data_folder = Path(os.getcwd())
    # file_to_open = data_folder / "ASH09_QC_LOG_BOOK.xlsm"
    # excel_file = pd.ExcelFile(file_to_open)

    excel_file_sht_er = pd.ExcelFile(excel_file_directory)
    df_reox1a_er = excel_file_sht_er.parse('REOX1A-ER', skiprows=9)                            # copy a sheet and paste into another sheet and skiprows 8
    df_reox1a_er['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL' in "Remarks" column 
    df_reox1a_er = df_reox1a_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_reox1a_er_sin = df_reox1a_er[df_reox1a_er["Layer"] == 'SiN']
    df_reox1a_er_sin = df_reox1a_er_sin.dropna()      # dropping rows where at least one element is missing
    df_reox1a_er_teos = df_reox1a_er[df_reox1a_er["Layer"] == 'TEOS']
    df_reox1a_er_teos = df_reox1a_er_teos.dropna()      # dropping rows where at least one element is missing

    # Display the dataframes in respective sheets
    # sht_reox1a_plot_barc.range('A25').options(index=False).value = df_reox1a_barc
    # sht_reox1a_plot_pr.range('A25').options(index=False).value = df_reox1a_pr
    # sht_reox1a_plot_teos.range('A25').options(index=False).value = df_reox1a_teos


    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for SiN ER & Unif PLot
    df_reox1a_er_sin_date = df_reox1a_er_sin["Date (MM/DD/YYYY)"]
    df_reox1a_er_sin_er = df_reox1a_er_sin["Etch Rate (A/Min)"]
    df_reox1a_er_sin_usl = df_reox1a_er_sin["USL"]
    df_reox1a_er_sin_lsl = df_reox1a_er_sin["LSL"]
    df_reox1a_er_sin_ucl = df_reox1a_er_sin["UCL"]
    df_reox1a_er_sin_lcl = df_reox1a_er_sin["LCL"]
    df_reox1a_er_sin_unif = df_reox1a_er_sin["% Uni"]
    df_reox1a_er_sin_unif_usl = df_reox1a_er_sin["% Uni USL"]
    df_reox1a_er_sin_unif_ucl = df_reox1a_er_sin["% Uni UCL"]
    df_reox1a_er_sin_remarks = df_reox1a_er_sin["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw SiN ER Plot (using Plotly) in Browser 
    draw_plotly_reox1a_er_sin_plot(
        x = date_formatter(df_reox1a_er_sin_date), 
        y1 = df_reox1a_er_sin_er,
        y2 = df_reox1a_er_sin_usl, 
        y3 = df_reox1a_er_sin_lsl,
        y4 = df_reox1a_er_sin_ucl,
        y5 = df_reox1a_er_sin_lcl,
        remarks = df_reox1a_er_sin_remarks
        )
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw SiN Unif Plot (using Plotly) in Browser     
    draw_plotly_reox1a_unif_sin_plot(
        x = date_formatter(df_reox1a_er_sin_date), 
        y1 = df_reox1a_er_sin_unif, 
        y2 = df_reox1a_er_sin_unif_usl,
        y3 = df_reox1a_er_sin_unif_ucl,
        remarks = df_reox1a_er_sin_remarks
        )

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Assigning variable to each param for TEOS ER & Unif PLot
    df_reox1a_er_teos_date = df_reox1a_er_teos["Date (MM/DD/YYYY)"]
    df_reox1a_er_teos_er = df_reox1a_er_teos["Etch Rate (A/Min)"]
    df_reox1a_er_teos_usl = df_reox1a_er_teos["USL"]
    df_reox1a_er_teos_lsl = df_reox1a_er_teos["LSL"]
    df_reox1a_er_teos_ucl = df_reox1a_er_teos["UCL"]
    df_reox1a_er_teos_lcl = df_reox1a_er_teos["LCL"]
    df_reox1a_er_teos_unif = df_reox1a_er_teos["% Uni"]
    df_reox1a_er_teos_unif_usl = df_reox1a_er_teos["% Uni USL"]
    df_reox1a_er_teos_unif_ucl = df_reox1a_er_teos["% Uni UCL"]
    df_reox1a_er_teos_remarks = df_reox1a_er_teos["Remarks"]
  
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw TEOS ER Plot (using Plotly) in Browser 
    draw_plotly_reox1a_er_teos_plot(
        x = date_formatter(df_reox1a_er_teos_date), 
        y1 = df_reox1a_er_teos_er,
        y2 = df_reox1a_er_teos_usl, 
        y3 = df_reox1a_er_teos_lsl,
        y4 = df_reox1a_er_teos_ucl,
        y5 = df_reox1a_er_teos_lcl,
        remarks = df_reox1a_er_teos_remarks
        )
   
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Draw TEOS Unif Plot (using Plotly) in Browser     
    draw_plotly_reox1a_unif_teos_plot(
        x = date_formatter(df_reox1a_er_teos_date), 
        y1 = df_reox1a_er_teos_unif, 
        y2 = df_reox1a_er_teos_unif_usl,
        y3 = df_reox1a_er_teos_unif_ucl,
        remarks = df_reox1a_er_teos_remarks
        )


#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
# User Defined Functions (UDFs)
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------
@xw.func
def hello(name):
    return "hello {0}".format(name)



