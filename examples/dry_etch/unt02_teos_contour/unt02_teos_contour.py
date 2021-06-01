import plotly as py
import plotly.graph_objs as go

import pandas as pd

# -----------------------------------------------------------------------
# input vars for MTRS1
filename_pre = 'TEOS_PRE_Contour_plot_UNT02_'
filename_er = 'TEOS_ER_Contour_plot_UNT02_'
titlename_pre = 'PRE-Thickness'
titlename_er = 'ER'

excel_file_directory = 'UNT02_Ch_B_QC_LOG_BOOK.xlsm'
# cols = ['site_1', 'site_2', 'site_3', 'site_4', 'site_5', 'site_6', 'site_7',
#         'site_8', 'site_9', 'site_10', 'site_11', 'site_12', 'site_13', 'site_14',
#         'site_15', 'site_16', 'site_17', 'site_18', 'site_19', 'site_20', 'site_21',
#         'site_22', 'site_23', 'site_24', 'site_25', 'site_26', 'site-27', 'site_28',
#         'site_29', 'site_30', 'site_31', 'site_32', 'site_33', 'site_34', 'site_35',
#         'site_36', 'site_37', 'site_38', 'site_39', 'site_40', 'site_41', 'site_42',
#         'site_43', 'site_44', 'site_45', 'site_46', 'site_47', 'site_48', 'site_49']

sht_name = 'ER-BARC,PR & TEOS'
sht_name_coords = 'teos_coords'
skiprows_ = 12
skiprows_coords = 0

coords_x = [0, 0, 0, 0, 0, -90, -45, 45, 90]
coords_y = [90, 45, 0, -45, -90, 0, 0, 0, 0]


# row_list_pre = [2802, 2811, 2820, 2829, 2832, 2841]
# row_list_post = [2803, 2812, 2821, 2830, 2833, 2842]
# row_list_er = [2804, 2813, 2822, 2831, 2834, 2843]

row_list_pre = [2797, 2806, 2815, 2818, 2827]
row_list_post = [2798, 2807, 2816, 2819, 2828]
row_list_er = [2799, 2808, 2817, 2820, 2829]

# ----------------------------------------------------------------------
# Read data from a csv
df = pd.read_excel(excel_file_directory, sheet_name= sht_name, skiprows= skiprows_)
# print(df)
df['Wafer Type'] = pd.Series(df['Wafer Type']).fillna(method='ffill')       # fill merged cells with same values
# df_coords = pd.read_excel(excel_file_directory, sheet_name= sht_name_coords, skiprows= skiprows_coords)
# print(df_coords)
df_teos = df[df['Wafer Type'] == 'ATZP91']
print(df_teos)
# df_teos.to_csv('out.csv')
print(df_teos.iloc[2].tolist())

x1 = coords_x
y1 = coords_y

# print(x1)
# print(y1)

def contour_plot(i, l, fname, tname):
    trace1 = go.Contour(
                x= x1,
                y= y1,
                z= df_teos.iloc[i].tolist(),
                contours=dict(
                            # coloring ='heatmap',
                            showlabels = True, # show labels on contours
                            labelfont = dict( # label font properties
                                size = 12,
                                color = 'white',
                               )
                        ),
                )
    

    layout = dict(
                title = 'Al ' + tname + ' Contour plot on MTRS1 for ' + 'RUN ' + str(l.index(i)+1),
                xaxis = dict(
                            title= 'x-axis (in mm)',
                            # type='linear',
                            range= [-100, 100]
                            ),
                yaxis = dict(
                            title= 'y-axis (in mm)',
                            # type='linear',
                            range= [-100, 100]
                            ),
                autosize= True,
                width=800, height=800,

                )

    data = [trace1]
    # fig = dict(data= data, layout= layout)
    fig = go.Figure(data= data, layout= layout)     # an `graph_objs` object created
    fig.add_trace(      # added a scatter plot for showing the markers onto the contour plot
            go.Scatter(
                    mode='markers',
                    x= x1,
                    y= y1,
                    opacity=0.8,
                    marker=dict(
                        color='#ffffff',
                        size=10,
                        line=dict(
                            color='#ffffff',
                            width= 0.5
                        ),
                        symbol = 'x'
                    ),
                    showlegend=False,
                    text= list(range(1,14))
            )        
        )
    fig.update_layout(      # added a unfilled circle from (-100, -100) to (100, 100) diagonally
        shapes=[
            # unfilled circle
            go.layout.Shape(
                type="circle",
                xref="x",
                yref="y",
                x0= -100,
                y0= -100,
                x1= 100,
                y1= 100,
                line_color="#212121",
            ),
        ]
    )
    py.offline.plot(fig, filename= fname + str(l.index(i)+1) + '.html')


# for i in row_list_pre:
#     contour_plot(i, row_list_pre, filename_pre, titlename_pre)

# for i in row_list_er:
#     contour_plot(i, row_list_er, filename_er, titlename_er)

# REFERENCE: 
# - https://help.plot.ly/documentation/python/shapes/
# - https://plot.ly/python/shapes/