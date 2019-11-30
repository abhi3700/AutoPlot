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
cp_plot_title = 'CP Plot for PRS01'  # title for CP plot
cp_plot_xlabel = 'Date'   # xaxis name for CP plot
cp_plot_ylabel = 'delta CP (no.s)'     # yaxis name for CP plot
cp_plot_html_file = 'PRS01_CP-Plot.html'   # HTML filename for CP plot
cp_plot_trace_count = 5    # no. of traces in CP plot
er_plot_title = 'ER Plot for PRS01'  # title for ER plot
er_plot_xlabel = 'Date'        # xaxis name for ER plot
er_plot_ylabel = 'ER (A/min)'   # yaxis name for ER plot
er_plot_html_file = 'PRS01_ER-Plot.html'   # HTML filename for ER plot
er_plot_trace_count = 5    # no. of traces in ER plot
sht_cp_name = 'QC DMIS Ref CP'
sht_er_name = 'QC DMIS Ref ER'
sht_cp_columns = ["Creationdate", "AverageValue", "UpperControlLimit", "LowerControlLimit", "UpperHoldValue", "LowerHoldValue", "Remarks"]
sht_er_columns = ["Creationdate", "AverageValue", "UpperControlLimit", "LowerControlLimit", "UpperHoldValue", "LowerHoldValue", "Remarks"]

excel_file_directory = "I:\\github_repos\\AutoPlot\\Examples\\wet_etch\\PRS01_Jan_19_to_Mar_19\\PRS01_Jan_19_to_Mar_19.xlsm"
# excel_file_directory = "\\\\vmfg\\VFD FILE SERVER\\SECTIONS\\DRY ETCH\\QC Log Book\\Final QC Log Book\\CNT_01_LOG_BOOK\\CNT01_Ch_A_QC_LOG_BOOK\\CNT01_Ch_A_QC_LOG_BOOK.xlsm"

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
"y3": LSL (y-axis) for CP Chart
"y4": UCL (y-axis) for CP Chart
"y5": LCL (y-axis) for CP Chart
"""
def draw_plotly_prs01_cp_plot(x, y1, y2, y3, y4, y5, remarks):
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
            name = 'UCL',
            mode = 'lines',
            line = dict(
                    color = cl_color,
                    width = 3)
    )

    data = [trace1, trace2, trace3, trace4, trace5]
    layout = dict(
            title = cp_plot_title,
            xaxis = dict(title= cp_plot_xlabel),
            yaxis = dict(title= cp_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= cp_plot_html_file, auto_open= False)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------
"""
"Description": This function plots ER Chart with `er_plot_trace_count` traces v/s Date.
"x": Date (x-axis) for ER Chart
"y1": ER (y-axis) for ER Chart
"y2": USL (y-axis) for ER Chart
"y3": LSL (y-axis) for ER Chart
"y4": UCL (y-axis) for ER Chart
"y5": LCL (y-axis) for ER Chart
"""
def draw_plotly_prs01_er_plot(x, y1, y2, y3, y4, y5, remarks):
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
            title = er_plot_title,
            xaxis = dict(title= er_plot_xlabel),
            yaxis = dict(title= er_plot_ylabel)
        )
    fig = dict(data= data, layout= layout)
    py.offline.plot(fig, filename= er_plot_html_file, auto_open= False)



#====================================================================================================================================================================
#####################################################################################################################################################################
def main():
    # wb = xw.Book.caller()
    # wb.sheets[0].range("A1").value = "Hello xlwings!"     # test code

    #****************************************************************************************************************************************************************
    # Define sheets
    # sht_prs01_cp = wb.sheets[sht_cp_name]
    # sht_prs01_er = wb.sheets[sht_er_name]
    # sht_prs01_plot_cp = wb.sheets['CP Plot']
    # sht_prs01_plot_er = wb.sheets['ER Plot']
    excel_file = pd.ExcelFile(excel_file_directory)

    #****************************************************************************************************************************************************************
    # Fetch Dataframe for CP Plot
    df_prs01_cp = excel_file.parse(sht_cp_name, skiprows=0)                            # copy a sheet and paste into another sheet and skiprows 9

    df_prs01_cp['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_prs01_cp = df_prs01_cp[sht_cp_columns]        # The final dataframe with required columns
    df_prs01_cp = df_prs01_cp.dropna()                                              # dropping rows where at least one element is missing
    # sht_prs01_plot_cp.range('A25').options(index=False).value = df_prs01_cp         # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param
    df_prs01_cp_date = df_prs01_cp[sht_cp_columns[0]]
    df_prs01_cp_delta_cp = df_prs01_cp[sht_cp_columns[1]]
    df_prs01_cp_usl = df_prs01_cp[sht_cp_columns[2]]
    df_prs01_cp_lsl = df_prs01_cp[sht_cp_columns[3]]
    df_prs01_cp_ucl = df_prs01_cp[sht_cp_columns[4]]
    df_prs01_cp_lcl = df_prs01_cp[sht_cp_columns[5]]
    df_prs01_cp_remarks = df_prs01_cp[sht_cp_columns[6]]

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw CP Plot (using Plotly) in Browser 
    draw_plotly_prs01_cp_plot(
        x = date_formatter(df_prs01_cp_date), 
        y1 = df_prs01_cp_delta_cp, 
        y2 = df_prs01_cp_usl, 
        y3 = df_prs01_cp_lsl, 
        y4 = df_prs01_cp_ucl, 
        y5 = df_prs01_cp_lcl, 
        remarks = df_prs01_cp_remarks
        )

    #****************************************************************************************************************************************************************
    df_prs01_er = excel_file.parse(sht_er_name, skiprows=0)                            # copy a sheet and paste into another sheet and skiprows 9
    
    df_prs01_er = df_prs01_er[sht_er_columns]             # The final Dataframe with 7 columns for plot: x-1, y-6
    df_prs01_er['Remarks'].fillna('NIL', inplace=True)        # replacing the empty cells with 'NIL'
    df_prs01_er = df_prs01_er.dropna()                                              # dropping rows where at least one element is misnitg
    # sht_prs01_plot_er.range('A28').options(index=False).value = df_prs01_er        # show the dataframe values into sheet- 'CP Plot'

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------    
    # Assigning variable to each param for ER & Unif PLot
    df_prs01_er_date = df_prs01_er[sht_cp_columns[0]]
    df_prs01_er_er = df_prs01_er[sht_cp_columns[1]]
    df_prs01_er_usl = df_prs01_er[sht_cp_columns[2]]
    df_prs01_er_lsl = df_prs01_er[sht_cp_columns[3]]
    df_prs01_er_ucl = df_prs01_er[sht_cp_columns[4]]
    df_prs01_er_lcl = df_prs01_er[sht_cp_columns[5]]
    df_prs01_er_remarks = df_prs01_er[sht_cp_columns[6]]
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Draw ER Plot (using Plotly) in Browser 
    draw_plotly_prs01_er_plot(
        x = date_formatter(df_prs01_er_date), 
        y1 = df_prs01_er_er,
        y2 = df_prs01_er_usl, 
        y3 = df_prs01_er_lsl,
        y4 = df_prs01_er_ucl,
        y5 = df_prs01_er_lcl,
        remarks = df_prs01_er_remarks
        )
   

if __name__ == '__main__':
	main()


