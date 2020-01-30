import plotly as py
import plotly.graph_objs as go

import pandas as pd

# -----------------------------------------------------------------------
# input vars for MTRS1
filename_pre = 'Al_PRE_Contour_plot_MTRS1_'
filename_er = 'Al_ER_Contour_plot_MTRS1_'
titlename_pre = 'PRE-Thickness'
titlename_er = 'ER'

excel_file_directory = 'Al_ER_and_unif.xlsx'
cols = ['site_1', 'site_2', 'site_3', 'site_4', 'site_5', 'site_6', 'site_7',
        'site_8', 'site_9', 'site_10', 'site_11', 'site_12', 'site_13', 'site_14',
        'site_15', 'site_16', 'site_17', 'site_18', 'site_19', 'site_20', 'site_21',
        'site_22', 'site_23', 'site_24', 'site_25', 'site_26', 'site-27', 'site_28',
        'site_29', 'site_30', 'site_31', 'site_32', 'site_33', 'site_34', 'site_35',
        'site_36', 'site_37', 'site_38', 'site_39', 'site_40', 'site_41', 'site_42',
        'site_43', 'site_44', 'site_45', 'site_46', 'site_47', 'site_48', 'site_49']

sht_name = '49 pt MTRS1 Al Rs'
skiprows_ = 11
coords_x = [0,0,30,0,-30,0,60,0,-60,0,90,0,-90]
coords_y = [0,-30,0,30,0,-60,0,60,0,-90,0,90,0]

row_list_pre = [2, 7, 12, 17, 22]
row_list_er = [4, 9, 14, 19, 24]

# ----------------------------------------------------------------------
# Read data from a csv
df = pd.read_excel(excel_file_directory, sheet_name= sht_name, skiprows= skiprows_)

df = df[cols]

# print(df)
# print(df.iloc[2].tolist())

x1 = coords_x
y1 = coords_y

def contour_plot(i, l, fname, tname):
    trace1 = go.Contour(
                x= x1,
                y= y1,
                z= df.iloc[i].tolist(),
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


for i in row_list_pre:
    contour_plot(i, row_list_pre, filename_pre, titlename_pre)

for i in row_list_er:
    contour_plot(i, row_list_er, filename_er, titlename_er)

# REFERENCE: 
# - https://help.plot.ly/documentation/python/shapes/
# - https://plot.ly/python/shapes/